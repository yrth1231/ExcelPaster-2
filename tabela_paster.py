"""
Tabela Paster - Windows program za lepljenje celija iz Excel tabele.

Hotkey-evi (rade samo dok je program upaljen, imaju prednost nad Windowsom):
    Ctrl+C            kopira CELIJU u kojoj se nalazis (bez lepljenja i pomeranja)
    Ctrl+V            nalepi CELIJU, predje na sledecu (desno) i odmah je kopira
    Ctrl+strelice     kretanje po tabeli (strelice bez Ctrl rade normalno)
    Ctrl+N            novi red, od prve kolone
    Ctrl+Shift+P      pauza / nastavak (privremeno vraca normalne Windows tastere)
"""
import csv
import datetime
import queue
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
PASTE_WAIT = 0.15   # koliko se ceka posle lepljenja (povecaj ako program kasni sa citanjem clipboarda)

COL_CURSOR = "#bcdcff"      # blago plava - trenutna celija
COL_CURSOR_BORDER = "#2b7de9"
COL_HDR = "#eeeeee"
COL_HDR_ACTIVE = "#cfe3ff"
COL_PINNED = "#fff3c4"      # pinovan prvi red


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
        root.geometry("540x300+20+20")
        root.minsize(280, 140)
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
        self.handles = {}
        self.colpos = []

        top = ttk.Frame(root)
        top.pack(fill="x", padx=2, pady=2)
        ttk.Button(top, text="Učitaj tabelu", command=self.load).pack(side="left")
        self.sheet_var = tk.StringVar()
        self.sheet_cb = ttk.Combobox(top, textvariable=self.sheet_var,
                                     state="readonly", width=10)
        self.sheet_cb.pack(side="left", padx=4)
        self.sheet_cb.bind("<<ComboboxSelected>>",
                           lambda e: self.load_sheet(self.sheet_var.get()))
        self.pin_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Pinuj prvi red", variable=self.pin_var,
                        command=self.refresh).pack(side="left", padx=4)
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
        root.after(5, self.poll)

    # ---------- hotkey-evi ----------
    def on_paste_key(self):
        self.q.put(("paste",))

    def on_copy_key(self):
        self.q.put(("copy",))

    def register_hotkeys(self):
        q = self.q
        pairs = [
            ("ctrl+left", lambda: q.put(("move", 0, -1))),
            ("ctrl+right", lambda: q.put(("move", 0, 1))),
            ("ctrl+up", lambda: q.put(("move", -1, 0))),
            ("ctrl+down", lambda: q.put(("move", 1, 0))),
            ("ctrl+n", lambda: q.put(("newrow",))),
            ("ctrl+v", self.on_paste_key),
            ("ctrl+c", self.on_copy_key),
        ]
        for key, fn in pairs:
            self.handles[key] = keyboard.add_hotkey(key, fn, suppress=True)

    def unregister_hotkeys(self):
        for h in list(self.handles.values()):
            try:
                keyboard.remove_hotkey(h)
            except Exception:
                pass
        self.handles = {}

    def toggle_pause(self):
        self.paused = not self.paused
        if self.paused:
            self.unregister_hotkeys()
            self.pause_btn.config(text="Nastavi")
        else:
            self.register_hotkeys()
            self.pause_btn.config(text="Pauza")
        self.update_status()

    # ---------- lepljenje ----------
    def set_clipboard(self, val):
        """Postavi clipboard i proveri da je zaista tu. Vraca True samo ako jeste."""
        want = val.replace("\r\n", "\n")
        for _ in range(15):
            try:
                pyperclip.copy(val)
                if pyperclip.paste().replace("\r\n", "\n") == want:
                    return True
            except Exception:
                pass
            time.sleep(0.02)
        return False

    def inject_combo(self, combo, letter):
        """Salje pravi Ctrl+<slovo> programu u fokusu. Nas hotkey se na
        trenutak skida da vestacki taster ne bi pokrenuo samog sebe."""
        h = self.handles.pop(combo, None)
        if h is not None:
            try:
                keyboard.remove_hotkey(h)
            except Exception:
                pass
        try:
            if keyboard.is_pressed("ctrl"):
                keyboard.send(letter)       # Ctrl je vec drzan
            else:
                keyboard.send(combo)
            time.sleep(PASTE_WAIT)          # daj programu vremena da procita clipboard
        finally:
            if h is not None and not self.paused:
                fn = self.on_paste_key if combo == "ctrl+v" else self.on_copy_key
                self.handles[combo] = keyboard.add_hotkey(combo, fn, suppress=True)

    def copy_cell(self):
        """Uvek upisuje trenutnu celiju u clipboard i proverava da je tu."""
        return self.set_clipboard(self.data[self.r][self.c])

    def do_copy(self):
        # Bez ucitane tabele - obican Ctrl+C
        if not self.data:
            self.inject_combo("ctrl+c", "c")
            return
        if self.copy_cell():
            self.update_status(prefix="✓ Kopirano   ")
        else:
            self.status.config(text="Clipboard je zauzet - pokušaj ponovo", fg="#b00020")

    def do_paste(self):
        # Bez ucitane tabele ne diramo clipboard - obican Ctrl+V
        if not self.data:
            self.inject_combo("ctrl+v", "v")
            return
        # 1) provera: u kojoj je celiji indikator (cita se tek sada, na pritisak)
        val = self.data[self.r][self.c]
        if val != "":
            # 2) kopiraj bas tu celiju (uvek upisuje, ne oslanja se na stari clipboard)
            if not self.copy_cell():
                self.status.config(text="Clipboard je zauzet - pokušaj ponovo", fg="#b00020")
                return                      # ne lepimo nista pogresno
            # 3) nalepi
            self.inject_combo("ctrl+v", "v")
        # 4) tek onda na sledecu celiju (desno); na kraju reda ostaje
        if self.c < self.ncols - 1:
            self.c += 1
        self.refresh()
        # unapred kopiraj novu celiju, da je sledeci Ctrl+V vec spreman
        if self.data[self.r][self.c] != "":
            self.copy_cell()

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
        # komande se obradjuju striktno redom, tako da Ctrl+V uvek
        # vidi tacnu, azurnu poziciju
        try:
            while True:
                cmd = self.q.get_nowait()
                if cmd[0] == "move":
                    self.move(cmd[1], cmd[2])
                elif cmd[0] == "newrow":
                    self.new_row()
                elif cmd[0] == "paste":
                    self.do_paste()
                elif cmd[0] == "copy":
                    self.do_copy()
                elif cmd[0] == "pause":
                    self.toggle_pause()
        except queue.Empty:
            pass
        self.root.after(5, self.poll)

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

    def min_row(self):
        return 1 if self.pin_var.get() else 0

    def top_h(self):
        return HDR_H + (ROW_H if self.pin_var.get() else 0)

    def on_wheel(self, e):
        if not self.data:
            return
        step = -3 if e.delta > 0 else 3
        self.sr = max(self.min_row(), min(self.nrows - 1, self.sr + step))
        self.draw()

    def on_click(self, e):
        if not self.data or e.y < HDR_H:
            return
        if self.pin_var.get() and e.y < HDR_H + ROW_H:
            i = 0
        else:
            i = max(self.sr, self.min_row()) + (e.y - self.top_h()) // ROW_H
        if i >= self.nrows:
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
        minr = self.min_row()
        vis = max(1, (h - self.top_h()) // ROW_H)
        if self.sr < minr:
            self.sr = minr
        if self.r >= minr:
            if self.r < self.sr:
                self.sr = self.r
            elif self.r >= self.sr + vis:
                self.sr = self.r - vis + 1
        if self.c < self.sc:
            self.sc = self.c
        else:
            while self.sc < self.c and sum(self.widths[self.sc:self.c + 1]) > w - RN_W:
                self.sc += 1

    def update_status(self, prefix=""):
        if self.paused:
            self.status.config(text="PAUZIRANO - normalni Windows tasteri", fg="#b00020")
            return
        if not self.data:
            self.status.config(text="Nema učitane tabele", fg="black")
            return
        val = self.data[self.r][self.c].replace("\n", " ")
        if len(val) > 40:
            val = val[:40] + "…"
        name = ""
        if self.pin_var.get() and self.r > 0:
            hn = self.data[0][self.c].strip()
            if hn:
                name = f" ({hn[:25]})"
        self.status.config(
            text=f"{prefix}Kolona {get_column_letter(self.c + 1)}{name}, red {self.r + 1}   |   {val}",
            fg="black")

    def refresh(self):
        if self.data:
            self.ensure_visible()
        self.update_status()
        self.draw()

    def draw_row(self, i, y, pinned_style=False):
        cv = self.canvas
        bg = COL_HDR_ACTIVE if i == self.r else COL_HDR
        cv.create_rectangle(0, y, RN_W, y + ROW_H, fill=bg, outline="#bbb")
        cv.create_text(RN_W / 2, y + ROW_H / 2, text=str(i + 1),
                       font=("Segoe UI", 9, "bold"))
        for j, x, cw in self.colpos:
            cur = (i == self.r and j == self.c)
            if cur:
                fill = COL_CURSOR
            elif pinned_style:
                fill = COL_PINNED
            else:
                fill = "white"
            cv.create_rectangle(x, y, x + cw, y + ROW_H, fill=fill,
                                outline=COL_CURSOR_BORDER if cur else "#ddd",
                                width=2 if cur else 1)
            txt = self.data[i][j].replace("\n", " ")
            n = max(1, (cw - 8) // CHAR_PX)
            if len(txt) > n:
                txt = txt[:n - 1] + "…"
            cv.create_text(x + 4, y + ROW_H / 2, text=txt, anchor="w",
                           font=("Segoe UI", 9, "bold" if pinned_style else "normal"))

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

        # zaglavlje kolona (slova)
        cv.create_rectangle(0, 0, RN_W, HDR_H, fill=COL_HDR, outline="#bbb")
        for j, x, cw in self.colpos:
            bg = COL_HDR_ACTIVE if j == self.c else COL_HDR
            cv.create_rectangle(x, 0, x + cw, HDR_H, fill=bg, outline="#bbb")
            cv.create_text(x + cw / 2, HDR_H / 2, text=get_column_letter(j + 1),
                           font=("Segoe UI", 9, "bold"))

        y = HDR_H
        pinned = self.pin_var.get()
        if pinned:
            self.draw_row(0, y, pinned_style=True)
            y += ROW_H
        i = max(self.sr, self.min_row())
        while i < self.nrows and y < h:
            self.draw_row(i, y)
            y += ROW_H
            i += 1

    def on_close(self):
        keyboard.unhook_all()   # vraca normalne Windows tastere
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
