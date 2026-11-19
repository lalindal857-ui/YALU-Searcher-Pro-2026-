import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import re, sys
import ctypes
import openpyxl

APP_TITLE = "YALU Searcher Pro"
APP_USER_MODEL_ID = "Lakshan.YALU.Searcher.Pro"

def set_windows_app_id():
    # Gives Windows a stable app identity so the custom EXE icon appears
    # correctly on the taskbar and in task switching.
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
        except Exception:
            pass

BG = "#06132f"
PANEL = "#081b42"
BLUE = "#087cff"
CYAN = "#00c9ff"
GREEN = "#00c853"
WHITE = "#f4f8ff"
MUTED = "#a9c7ed"

def asset_path(name):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name

class YaluSearcher:
    def __init__(self, root):
        self.root = root
        self.root.title("YALU Searcher")
        self.root.geometry("283x114")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=BG)
        self.rows, self.col, self.last_copy = [], {}, ""
        self._drag_offset = (0, 0)
        try:
            icon_file = str(asset_path("yalu_searcher.ico"))
            self.root.iconbitmap(default=icon_file)
            self.root.wm_iconbitmap(icon_file)
        except Exception:
            pass
        self._ui()
        self.root.bind("<Return>", lambda e: self.search_copy())
        self.root.bind("<Control-o>", lambda e: self.open_excel())

    def _drag_start(self, event):
        self._drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())

    def _drag_move(self, event):
        x = event.x_root - self._drag_offset[0]
        y = event.y_root - self._drag_offset[1]
        self.root.geometry(f"+{x}+{y}")

    def _drag_bind(self, widget):
        widget.bind("<ButtonPress-1>", self._drag_start, add="+")
        widget.bind("<B1-Motion>", self._drag_move, add="+")

    def _ui(self):
        # Small floating window with visible cost center and author credit.
        outer = tk.Frame(self.root, bg=BG, highlightbackground="#168bff", highlightthickness=1)
        outer.pack(fill="both", expand=True, padx=1, pady=1)

        header = tk.Frame(outer, bg="#040d22", height=23)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="✦  YALU SEARCHER", bg="#040d22", fg=CYAN,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=9)
        tk.Label(header, text="JE", bg="#0756d8", fg=WHITE,
                 font=("Segoe UI", 8, "bold"), padx=6, pady=1).pack(side="right", padx=7, pady=2)
        self._drag_bind(header)
        for w in header.winfo_children():
            self._drag_bind(w)

        search = tk.Frame(outer, bg=BG)
        search.pack(fill="x", padx=7, pady=(7, 3))
        tk.Button(search, text="📂", command=self.open_excel, bg="#126bdb", fg=WHITE,
                  activebackground="#0756d8", activeforeground=WHITE, relief="flat", bd=0,
                  width=3, font=("Segoe UI", 10, "bold"), cursor="hand2").pack(side="left", ipady=2)
        self.entry = tk.Entry(search, justify="center", bg="#102957", fg=WHITE,
                              insertbackground=WHITE, relief="flat", bd=0,
                              font=("Segoe UI", 10, "bold"))
        self.entry.pack(side="left", fill="x", expand=True, padx=6, ipady=4)
        self.entry.insert(0, "Entry No.")
        self.entry.bind("<FocusIn>", self._clear_hint)
        tk.Button(search, text="COPY", command=self.search_copy, bg=GREEN, fg=WHITE,
                  activebackground="#00a844", activeforeground=WHITE, relief="flat", bd=0,
                  width=6, font=("Segoe UI", 9, "bold"), cursor="hand2").pack(side="right", ipady=4)

        info = tk.Frame(outer, bg="#091a3a")
        info.pack(fill="x", padx=7, pady=(1, 3))
        self.status_var = tk.StringVar(value="Select Excel file")
        self.cost_var = tk.StringVar(value="Cost Center: —")
        tk.Label(info, textvariable=self.status_var, bg="#091a3a", fg="#43dfa0",
                 font=("Segoe UI", 7, "bold"), anchor="w").pack(fill="x", padx=7, pady=(3, 0))
        tk.Label(info, textvariable=self.cost_var, bg="#091a3a", fg="#c8dcff",
                 font=("Segoe UI", 7), anchor="w").pack(fill="x", padx=7, pady=(0, 3))

        footer = tk.Frame(outer, bg="#040d22", height=18)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(footer, text="● READY", bg="#040d22", fg="#43dfa0",
                 font=("Segoe UI", 7, "bold")).pack(side="left", padx=8)
        tk.Label(footer, text="Prepared By Lakshan", bg="#040d22", fg=MUTED,
                 font=("Segoe UI", 7, "bold")).pack(side="right", padx=8)

        self.root.bind("<Control-o>", lambda e: self.open_excel())
        self.root.bind("<Control-l>", lambda e: self.show_last())

    def minimize(self):
        # Minimize the floating window; it can be restored from the taskbar.
        self.root.iconify()

    def _clear_hint(self, _):
        if self.entry.get() == "Entry No.":
            self.entry.delete(0, "end")

    @staticmethod
    def clean(v):
        return "" if v is None else str(v).strip()

    @staticmethod
    def is_header(v):
        return bool(re.match(r"^JE-", v.strip(), re.I))

    def open_excel(self):
        path = filedialog.askopenfilename(
            title="Select Journal Entries report",
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if "Journal Entries" not in wb.sheetnames:
                raise ValueError("Sheet 'Journal Entries' not found.")
            ws = wb["Journal Entries"]
            headers = [self.clean(v) for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
            required = ["Entry No. / Line", "Posting Date / Cost Center", "Description"]
            missing = [h for h in required if h not in headers]
            if missing:
                raise ValueError("Missing columns: " + ", ".join(missing))
            self.col = {h: headers.index(h) for h in required}
            self.rows = list(ws.iter_rows(min_row=2, values_only=True))
            wb.close()
            self.status_var.set(f"Loaded {len(self.rows):,} rows")
            self.entry.focus_set()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))

    def search_copy(self):
        q = self.entry.get().strip()
        if not self.rows:
            messagebox.showinfo(APP_TITLE, "Open an Excel report first.")
            return
        if not q or q == "Entry No.":
            return
        ei = self.col["Entry No. / Line"]
        ci = self.col["Posting Date / Cost Center"]
        di = self.col["Description"]
        start, full_entry = None, ""
        for i, r in enumerate(self.rows):
            val = self.clean(r[ei])
            if self.is_header(val) and (
                val.casefold() == q.casefold()
                or val.casefold().endswith("-" + q.casefold())
                or val.casefold().endswith(q.casefold())
            ):
                start, full_entry = i, val
                break
        if start is None:
            self.status_var.set("Not found")
            self.cost_var.set("Cost Center: —")
            return
        output, centers = [], []
        for r in self.rows[start + 1:]:
            raw = self.clean(r[ei])
            if self.is_header(raw):
                break
            if not re.fullmatch(r"\d+", raw):
                continue
            desc, center = self.clean(r[di]), self.clean(r[ci])
            # Put the full Journal Entry number at the END of each description.
            if desc:
                output.append(f"{desc} ({full_entry})")
            if center and center not in centers:
                centers.append(center)
        if not output:
            self.status_var.set("No detail lines")
            return
        self.last_copy = "\n".join(output)
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_copy)
        self.root.update()
        self.cost_var.set("Cost Center: " + (", ".join(centers) if centers else "—"))
        self.status_var.set("Copied ✓")
        self.root.bell()
        self.show_prepared_notification()

    def show_prepared_notification(self):
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(bg=BLUE)
        width, height = 230, 62
        x = self.root.winfo_screenwidth() - width - 22
        y = self.root.winfo_screenheight() - height - 70
        toast.geometry(f"{width}x{height}+{x}+{y}")
        panel = tk.Frame(toast, bg="#071a3c", highlightbackground=CYAN, highlightthickness=1)
        panel.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(panel, text="✓  COPIED SUCCESSFULLY", bg="#071a3c", fg="#36e58e",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(7, 1))
        tk.Label(panel, text="Prepared By Lakshan", bg="#071a3c", fg=WHITE,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10)
        toast.after(2500, toast.destroy)

    def show_last(self):
        if self.last_copy:
            messagebox.showinfo("Copied Description", self.last_copy)
        else:
            messagebox.showinfo(APP_TITLE, "No search result yet.")

if __name__ == "__main__":
    set_windows_app_id()
    root = tk.Tk()
    YaluSearcher(root)
    root.mainloop()
