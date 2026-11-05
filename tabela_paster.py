"""
Tabela Paster - Windows program za lepljenje celija iz Excel tabele.

Hotkey-evi (rade samo dok je program upaljen):
    Ctrl+B            nalepi trenutnu celiju u program u kome si
    Ctrl+strelice     kretanje po tabeli (strelice bez Ctrl rade normalno)
    Ctrl+N            novi red, od prve kolone
    Ctrl+Shift+P      pauza / nastavak (privremeno vraca normalne Windows tastere)
"""
import csv
import datetime
import queue
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import keyboard
import pyperclip
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import fills, fonts, borders, alignment, protection


# Zaobilazi gresku u nekim xlsx fajlovima:
# "__init__() got an unexpected keyword argument 'extLst'"
def _ignore_extlst(cls):
    orig = cls.__init__

    def init(self, *args, **kwargs):
        kwargs.pop("extLst", None)
        orig(self, *args, **kwargs)

    cls.__init__ = init


for _cls in (fills.PatternFill, fills.GradientFill, fonts.Font, borders.Border,
             borders.Side, alignment.Alignment, protection.Protection):
    _ignore_extlst(_cls)

ROW_H = 22
HDR_H = 22
RN_W = 44
CHAR_PX = 7

COL_CURSOR = "#bcdcff"      # blago plava - trenutna celija
COL_CURSOR_BORDER = "#2b7de9"
COL_HDR = "#eeeeee"
COL_HDR_ACTIVE = "#cfe3ff"


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%d.%m.%Y")
        return v.strftime("%d.%m.%Y %H:%M")
    if isinstance(v, datetime.date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


class App:
    def __init__(self, root):
        self.root = root
        root.title("Tabela Paster")
        root.geometry("460x280+20+20")
        root.minsize(260, 140)
        root.attributes("-topmost", True)

        self.q = queue.Queue()
        self.data = []
        self.widths = []
        self.nrows = 0
        self.ncols = 0
        self.r = 0
        self.c = 0
        self.sr = 0
        self.sc = 0
        self.path = None
        self.paused = False
        self.handles = []
        self.colpos = []

        top = ttk.Frame(root)
        top.pack(fill="x", padx=2, pady=2)
        ttk.Button(top, text="Učitaj tabelu", command=self.load).pack(side="left")
        self.sheet_var = tk.StringVar()
        self.sheet_cb = ttk.Combobox(top, textvariable=self.sheet_var,
                                     state="readonly", width=14)
        self.sheet_cb.pack(side="left", padx=4)
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self.load_sheet(self.sheet_var.get()))
        self.pause_btn = ttk.Button(top, text="Pauza", command=self.toggle_pause)
        self.pause_btn.pack(side="right")

        self.status = tk.Label(root, anchor="w", font=("Segoe UI", 9, "bold"),
                               text="Nema učitane tabele")
        self.status.pack(fill="x", padx=4)

        self.canvas = tk.Canvas(root, bg="white", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.refresh())
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<MouseWheel>", self.on_wheel)

        self.register_hotkeys()
        keyboard.add_hotkey("ctrl+shift+p", lambda: self.q.put(("pause",)), suppress=True)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(10, self.poll)

    # ---------- hotkey-evi ----------
    def register_hotkeys(self):
        q = self.q
        pairs = [
            ("ctrl+left", lambda: q.put(("move", 0, -1))),
            ("ctrl+right", lambda: q.put(("move", 0, 1))),
            ("ctrl+up", lambda: q.put(("move", -1, 0))),
            ("ctrl+down", lambda: q.put(("move", 1, 0))),
            ("ctrl+n", lambda: q.put(("newrow",))),
            ("ctrl+b", lambda: threading.Thread(target=self.do_paste, daemon=True).start()),
        ]
        for key, fn in pairs:
            self.handles.append(keyboard.add_hotkey(key, fn, suppress=True))

    def unregister_hotkeys(self):
        for h in self.handles:
            try:
                keyboard.remove_hotkey(h)
            except Exception:
                pass
        self.handles = []

    def toggle_pause(self):
        self.paused = not self.paused
        if self.paused:
            self.unregister_hotkeys()
            self.pause_btn.config(text="Nastavi")
        else:
            self.register_hotkeys()
            self.pause_btn.config(text="Pauza")
        self.update_status()

    def do_paste(self):
        if not self.data:
            return
        try:
            val = self.data[self.r][self.c]
        except IndexError:
            return
        pyperclip.copy(val)
        time.sleep(0.03)
        # Ctrl je vec drzan dok pritiskas Ctrl+B, pa saljemo samo V
        if keyboard.is_pressed("ctrl"):
            keyboard.send("v")
        else:
            keyboard.send("ctrl+v")

    # ---------- učitavanje ----------
    def load(self):
        path = filedialog.askopenfilename(
            title="Izaberi tabelu",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("Svi fajlovi", "*.*")])
        if not path:
            return
        self.path = path
        try:
            if path.lower().endswith(".csv"):
                self.sheet_cb["values"] = []
                self.sheet_var.set("")
                self.load_csv(path)
            else:
                wb = load_workbook(path, read_only=True, data_only=True)
                names = wb.sheetnames
                wb.close()
                self.sheet_cb["values"] = names
                self.sheet_var.set(names[0])
                self.load_sheet(names[0])
        except Exception as e:
            messagebox.showerror("Greška", f"Ne mogu da učitam fajl:\n{e}")

    def load_csv(self, path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            rows = [list(r) for r in csv.reader(f, dialect)]
        self.set_data(rows)

    def load_sheet(self, name):
        try:
            wb = load_workbook(self.path, read_only=True, data_only=True)
            ws = wb[name]
            rows = [[fmt(v) for v in row] for row in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception as e:
            messagebox.showerror("Greška", f"Ne mogu da učitam list:\n{e}")
            return
        self.set_data(rows)

    def set_data(self, rows):
        while rows and not any(str(x).strip() for x in rows[-1]):
            rows.pop()
        ncols = max((len(r) for r in rows), default=0)
        rows = [[str(x) for x in r] + [""] * (ncols - len(r)) for r in rows]
        self.data = rows
        self.nrows = len(rows)
        self.ncols = ncols
        self.widths = []
        for j in range(ncols):
            longest = len(get_column_letter(j + 1))
            for i in range(min(self.nrows, 300)):
                longest = max(longest, len(rows[i][j]))
            self.widths.append(max(4, min(longest, 30)) * CHAR_PX + 12)
        self.r = self.c = self.sr = self.sc = 0
        self.refresh()

    # ---------- kretanje ----------
    def poll(self):
        try:
            while True:
                cmd = self.q.get_nowait()
                if cmd[0] == "move":
                    self.move(cmd[1], cmd[2])
                elif cmd[0] == "newrow":
                    self.new_row()
                elif cmd[0] == "pause":
                    self.toggle_pause()
        except queue.Empty:
            pass
        self.root.after(10, self.poll)

    def move(self, dr, dc):
        if not self.data:
            return
        self.r = max(0, min(self.nrows - 1, self.r + dr))
        self.c = max(0, min(self.ncols - 1, self.c + dc))
        self.refresh()

    def new_row(self):
        if not self.data:
            return
        self.c = 0
        self.r = min(self.nrows - 1, self.r + 1)
        self.refresh()

    def on_wheel(self, e):
        if not self.data:
            return
        self.sr = max(0, min(self.nrows - 1, self.sr - (1 if e.delta > 0 else -1) * 3))
        self.draw()

    def on_click(self, e):
        if not self.data:
            return
        i = self.sr + (e.y - HDR_H) // ROW_H
        if e.y < HDR_H or i >= self.nrows:
            return
        for j, x, w in self.colpos:
            if x <= e.x < x + w:
                self.r, self.c = i, j
                self.refresh()
                return

    # ---------- prikaz ----------
    def ensure_visible(self):
        h = self.canvas.winfo_height()
        w = self.canvas.winfo_width()
        vis = max(1, (h - HDR_H) // ROW_H)
        if self.r < self.sr:
            self.sr = self.r
        elif self.r >= self.sr + vis:
            self.sr = self.r - vis + 1
        if self.c < self.sc:
            self.sc = self.c
        else:
            while self.sc < self.c and sum(self.widths[self.sc:self.c + 1]) > w - RN_W:
                self.sc += 1

    def update_status(self):
        if self.paused:
            self.status.config(text="PAUZIRANO - normalni Windows tasteri", fg="#b00020")
            return
        if not self.data:
            self.status.config(text="Nema učitane tabele", fg="black")
            return
        val = self.data[self.r][self.c].replace("\n", " ")
        if len(val) > 40:
            val = val[:40] + "…"
        self.status.config(
            text=f"Kolona {get_column_letter(self.c + 1)}, red {self.r + 1}   |   {val}",
            fg="black")

    def refresh(self):
        if self.data:
            self.ensure_visible()
        self.update_status()
        self.draw()

    def draw(self):
        cv = self.canvas
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if not self.data:
            cv.create_text(w / 2, h / 2, text="Klikni „Učitaj tabelu“", fill="#888")
            return

        self.colpos = []
        x = RN_W
        j = self.sc
        while j < self.ncols and x < w:
            self.colpos.append((j, x, self.widths[j]))
            x += self.widths[j]
            j += 1

        nvis = (h - HDR_H) // ROW_H + 1
        last = min(self.nrows, self.sr + nvis)

        # zaglavlje kolona
        cv.create_rectangle(0, 0, RN_W, HDR_H, fill=COL_HDR, outline="#bbb")
        for j, x, cw in self.colpos:
            bg = COL_HDR_ACTIVE if j == self.c else COL_HDR
            cv.create_rectangle(x, 0, x + cw, HDR_H, fill=bg, outline="#bbb")
            cv.create_text(x + cw / 2, HDR_H / 2, text=get_column_letter(j + 1),
                           font=("Segoe UI", 9, "bold"))

        for i in range(self.sr, last):
            y = HDR_H + (i - self.sr) * ROW_H
            bg = COL_HDR_ACTIVE if i == self.r else COL_HDR
            cv.create_rectangle(0, y, RN_W, y + ROW_H, fill=bg, outline="#bbb")
            cv.create_text(RN_W / 2, y + ROW_H / 2, text=str(i + 1),
                           font=("Segoe UI", 9, "bold"))
            for j, x, cw in self.colpos:
                cur = (i == self.r and j == self.c)
                cv.create_rectangle(x, y, x + cw, y + ROW_H,
                                    fill=COL_CURSOR if cur else "white",
                                    outline=COL_CURSOR_BORDER if cur else "#ddd",
                                    width=2 if cur else 1)
                txt = self.data[i][j].replace("\n", " ")
                n = max(1, (cw - 8) // CHAR_PX)
                if len(txt) > n:
                    txt = txt[:n - 1] + "…"
                cv.create_text(x + 4, y + ROW_H / 2, text=txt, anchor="w",
                               font=("Segoe UI", 9))

    def on_close(self):
        keyboard.unhook_all()   # vraca normalne Windows tastere
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
