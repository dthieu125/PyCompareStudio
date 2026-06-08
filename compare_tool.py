import csv
import difflib
import filecmp
import gzip
import hashlib
import argparse
import json
import mimetypes
import os
import re
import shutil
import sys
import tarfile
import tempfile
import threading
import keyword
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    HORIZONTAL,
    LEFT,
    RIGHT,
    VERTICAL,
    W,
    X,
    Y,
    BooleanVar,
    IntVar,
    StringVar,
    Tk,
    filedialog,
    messagebox,
    colorchooser,
    ttk,
)
from tkinter.scrolledtext import ScrolledText


APP_TITLE = "PyCompare Studio"


TEXT_EXTENSIONS = {
    ".bat",
    ".c",
    ".cfg",
    ".conf",
    ".cpp",
    ".cs",
    ".css",
    ".go",
    ".h",
    ".hpp",
    ".htm",
    ".html",
    ".ini",
    ".java",
    ".js",
    ".json",
    ".log",
    ".md",
    ".php",
    ".ps1",
    ".py",
    ".pyw",
    ".rb",
    ".rs",
    ".sh",
    ".sql",
    ".ts",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}
TABLE_EXTENSIONS = {".csv", ".xlsx"}
PICTURE_EXTENSIONS = {".bmp", ".gif", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".webp"}
STATE_DIR = Path(os.environ.get("APPDATA") or Path.home() / ".config") / "PyCompareStudio"
LEFT_STATE_FILE = STATE_DIR / "selected_left.json"
SETTINGS_FILE = STATE_DIR / "settings.json"
DEFAULT_SETTINGS = {
    "match_color": "#eaf7ea",
    "diff_color": "#fff0cf",
    "only_left_color": "#dcecff",
    "only_right_color": "#ffe2e2",
    "text_delete_color": "#ffe2e2",
    "text_insert_color": "#dcecff",
    "text_diff_color": "#fff0cf",
    "startup_windows": False,
}


try:
    import openpyxl
except ImportError:  # Optional dependency.
    openpyxl = None

try:
    from PIL import Image, ImageChops, ImageTk
except ImportError:  # Optional dependency.
    Image = ImageChops = ImageTk = None

try:
    import py7zr
except ImportError:  # Optional dependency.
    py7zr = None


@dataclass
class FolderItem:
    rel_path: str
    status: str
    left: Path | None
    right: Path | None
    size_left: int | None = None
    size_right: int | None = None


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def normalize_text(text: str, ignore_ws: bool, ignore_blank: bool, ignore_comments: bool) -> list[str]:
    lines = text.splitlines()
    normalized: list[str] = []
    block_comment = False
    for line in lines:
        working = line
        if ignore_comments:
            if block_comment:
                end = working.find("*/")
                if end >= 0:
                    working = working[end + 2 :]
                    block_comment = False
                else:
                    continue
            while "/*" in working:
                start = working.find("/*")
                end = working.find("*/", start + 2)
                if end >= 0:
                    working = working[:start] + working[end + 2 :]
                else:
                    working = working[:start]
                    block_comment = True
                    break
            working = re.sub(r"#.*$", "", working)
            working = re.sub(r"//.*$", "", working)
        if ignore_ws:
            working = re.sub(r"\s+", " ", working).strip()
        if ignore_blank and not working.strip():
            continue
        normalized.append(working)
    return normalized


def read_text_file(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1258", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def write_text_file(path: Path, value: str) -> None:
    path.write_text(value, encoding="utf-8", newline="")


def load_settings() -> dict:
    if not SETTINGS_FILE.exists():
        return DEFAULT_SETTINGS.copy()
    try:
        loaded = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return DEFAULT_SETTINGS.copy()
    settings = DEFAULT_SETTINGS.copy()
    settings.update({key: value for key, value in loaded.items() if key in settings})
    return settings


def save_settings(settings: dict) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def app_launch_command() -> str:
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    python = Path(sys.executable)
    if os.name == "nt":
        pythonw = python.with_name("pythonw.exe")
        python = pythonw if pythonw.exists() else python
    return f'"{python}" "{Path(__file__).resolve()}"'


def set_windows_startup(enabled: bool) -> None:
    if os.name != "nt":
        return
    import winreg

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
        if enabled:
            winreg.SetValueEx(key, "PyCompareStudio", 0, winreg.REG_SZ, app_launch_command())
        else:
            try:
                winreg.DeleteValue(key, "PyCompareStudio")
            except FileNotFoundError:
                pass


def is_windows_startup_enabled() -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg

        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.QueryValueEx(key, "PyCompareStudio")
            return True
    except OSError:
        return False


def notify(title: str, message: str) -> None:
    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def notify_error(title: str, message: str) -> None:
    root = Tk()
    root.withdraw()
    messagebox.showerror(title, message)
    root.destroy()


def save_left_selection(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(path)
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    LEFT_STATE_FILE.write_text(
        json.dumps({"path": str(path), "name": path.name}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    update_windows_compare_label(path.name)


def load_left_selection() -> Path:
    if not LEFT_STATE_FILE.exists():
        raise RuntimeError("No left file selected. Use 'Select Left File for Compare' first.")
    data = json.loads(LEFT_STATE_FILE.read_text(encoding="utf-8"))
    path = Path(data["path"])
    if not path.exists():
        raise RuntimeError(f"Selected left file no longer exists: {path}")
    return path


def update_windows_compare_label(left_name: str) -> None:
    if os.name != "nt":
        return
    try:
        import winreg

        key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PyCompareStudio.CompareToLeft"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, f"Compare to {left_name}")
    except OSError:
        return


def run_context_action(args: argparse.Namespace) -> bool:
    try:
        if args.select_left:
            path = Path(args.select_left)
            save_left_selection(path)
            notify(APP_TITLE, f"Selected left file:\n{path}")
            return True
        if args.compare_to_left:
            left = load_left_selection()
            right = Path(args.compare_to_left)
            if not right.exists():
                raise FileNotFoundError(right)
            app = App(str(left), str(right))
            app.mainloop()
            return True
        if args.compare_selected:
            selected = normalize_selected_paths(args.compare_selected)
            if len(selected) != 2:
                raise RuntimeError("Please select exactly 2 files or folders to compare.")
            missing = [str(path) for path in selected if not path.exists()]
            if missing:
                raise FileNotFoundError(", ".join(missing))
            app = App(str(selected[0]), str(selected[1]))
            app.mainloop()
            return True
    except Exception as exc:
        notify_error(APP_TITLE, str(exc))
        return True
    return False


def normalize_selected_paths(values: list[str]) -> list[Path]:
    candidates = [value for value in values if value]
    if len(candidates) == 1:
        blob = candidates[0]
        if "\n" in blob:
            candidates = [part for part in blob.splitlines() if part.strip()]
        elif ";" in blob:
            candidates = [part for part in blob.split(";") if part.strip()]
        else:
            try:
                import shlex

                split_values = shlex.split(blob, posix=False)
                if len(split_values) > 1:
                    candidates = [part.strip('"') for part in split_values]
            except ValueError:
                pass
    paths = [Path(value.strip('"')) for value in candidates]
    if len(paths) <= 2:
        return paths
    rebuilt: list[Path] = []
    index = 0
    while index < len(candidates):
        current = candidates[index].strip('"')
        probe = Path(current)
        while not probe.exists() and index + 1 < len(candidates):
            index += 1
            current += " " + candidates[index].strip('"')
            probe = Path(current)
        rebuilt.append(probe)
        index += 1
    return rebuilt


def open_virtual_folder(source: Path) -> tempfile.TemporaryDirectory | None:
    suffix = source.suffix.lower()
    temp_dir = tempfile.TemporaryDirectory(prefix="pycompare_")
    target = Path(temp_dir.name)
    try:
        if suffix == ".zip":
            with zipfile.ZipFile(source) as archive:
                archive.extractall(target)
        elif suffix in {".tar", ".tgz", ".gz", ".gzip", ".bz2", ".xz"}:
            if suffix in {".gz", ".gzip"} and not source.name.endswith((".tar.gz", ".tgz")):
                output = target / source.with_suffix("").name
                with gzip.open(source, "rb") as src, output.open("wb") as dst:
                    shutil.copyfileobj(src, dst)
            else:
                with tarfile.open(source) as archive:
                    archive.extractall(target)
        elif suffix == ".7z" and py7zr is not None:
            with py7zr.SevenZipFile(source, mode="r") as archive:
                archive.extractall(target)
        else:
            temp_dir.cleanup()
            return None
        return temp_dir
    except Exception:
        temp_dir.cleanup()
        raise


def collect_files(root: Path) -> dict[str, Path]:
    result: dict[str, Path] = {}
    for current, _, files in os.walk(root):
        for name in files:
            full_path = Path(current) / name
            result[str(full_path.relative_to(root)).replace("\\", "/")] = full_path
    return result


def compare_folders(left: Path, right: Path, deep: bool) -> list[FolderItem]:
    left_files = collect_files(left)
    right_files = collect_files(right)
    all_paths = sorted(set(left_files) | set(right_files), key=str.lower)
    items: list[FolderItem] = []
    for rel_path in all_paths:
        left_path = left_files.get(rel_path)
        right_path = right_files.get(rel_path)
        if left_path and not right_path:
            items.append(FolderItem(rel_path, "Only left", left_path, None, left_path.stat().st_size, None))
        elif right_path and not left_path:
            items.append(FolderItem(rel_path, "Only right", None, right_path, None, right_path.stat().st_size))
        elif left_path and right_path:
            left_size = left_path.stat().st_size
            right_size = right_path.stat().st_size
            if left_size != right_size:
                status = "Different"
            elif deep:
                status = "Match" if sha256_file(left_path) == sha256_file(right_path) else "Different"
            else:
                status = "Match" if filecmp.cmp(left_path, right_path, shallow=True) else "Different"
            items.append(FolderItem(rel_path, status, left_path, right_path, left_size, right_size))
    return items


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


class PathPicker(ttk.Frame):
    def __init__(self, parent, label: str, mode: str = "file", allow_archive: bool = False):
        super().__init__(parent)
        self.mode = mode
        self.allow_archive = allow_archive
        self.value = StringVar()
        ttk.Label(self, text=label, width=10).pack(side=LEFT)
        ttk.Entry(self, textvariable=self.value).pack(side=LEFT, fill=X, expand=True, padx=4)
        ttk.Button(self, text="Browse", command=self.browse).pack(side=RIGHT)
        if allow_archive:
            ttk.Button(self, text="Archive", command=self.browse_file).pack(side=RIGHT, padx=(0, 4))

    def browse(self) -> None:
        if self.mode == "folder":
            selected = filedialog.askdirectory()
        else:
            selected = filedialog.askopenfilename()
        if selected:
            self.value.set(selected)

    def browse_file(self) -> None:
        selected = filedialog.askopenfilename(
            filetypes=[
                ("Archives", "*.zip *.tar *.tgz *.tar.gz *.gz *.gzip *.7z"),
                ("All files", "*.*"),
            ]
        )
        if selected:
            self.value.set(selected)

    def path(self) -> Path | None:
        raw = self.value.get().strip()
        return Path(raw) if raw else None


class FolderCompareTab(ttk.Frame):
    def __init__(self, parent, settings: dict):
        super().__init__(parent)
        self.settings = settings
        self.left_temp = None
        self.right_temp = None
        self.items: list[FolderItem] = []
        self.deep = BooleanVar(value=True)
        self.status = StringVar(value="Ready")
        self._build()

    def _build(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=X, padx=8, pady=8)
        self.left = PathPicker(top, "Left", "folder", allow_archive=True)
        self.left.pack(fill=X, pady=2)
        self.right = PathPicker(top, "Right", "folder", allow_archive=True)
        self.right.pack(fill=X, pady=2)

        controls = ttk.Frame(self)
        controls.pack(fill=X, padx=8)
        ttk.Checkbutton(controls, text="Deep content compare", variable=self.deep).pack(side=LEFT)
        ttk.Button(controls, text="Compare", command=self.start_compare).pack(side=LEFT, padx=4)
        ttk.Button(controls, text="Copy left -> right", command=lambda: self.sync("ltr")).pack(side=LEFT, padx=4)
        ttk.Button(controls, text="Copy right -> left", command=lambda: self.sync("rtl")).pack(side=LEFT, padx=4)
        ttk.Button(controls, text="Mirror left -> right", command=lambda: self.sync("mirror_ltr")).pack(side=LEFT, padx=4)
        ttk.Button(controls, text="Mirror right -> left", command=lambda: self.sync("mirror_rtl")).pack(side=LEFT, padx=4)

        columns = ("path", "status", "left_size", "right_size")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("path", text="Relative path")
        self.tree.heading("status", text="Status")
        self.tree.heading("left_size", text="Left size")
        self.tree.heading("right_size", text="Right size")
        self.tree.column("path", width=520)
        self.tree.column("status", width=110, anchor=W)
        self.tree.pack(fill=BOTH, expand=True, padx=8, pady=8)
        self.apply_settings()
        ttk.Label(self, textvariable=self.status).pack(fill=X, padx=8, pady=(0, 8))

    def apply_settings(self) -> None:
        self.tree.tag_configure("Match", background=self.settings["match_color"])
        self.tree.tag_configure("Different", background=self.settings["diff_color"])
        self.tree.tag_configure("Only left", background=self.settings["only_left_color"])
        self.tree.tag_configure("Only right", background=self.settings["only_right_color"])

    def resolve_source(self, picker: PathPicker) -> tuple[Path, tempfile.TemporaryDirectory | None]:
        source = picker.path()
        if not source:
            raise ValueError("Please choose both paths.")
        if source.is_dir():
            return source, None
        extracted = open_virtual_folder(source)
        if extracted is None:
            raise ValueError(f"Unsupported folder source: {source}")
        return Path(extracted.name), extracted

    def start_compare(self) -> None:
        self.status.set("Comparing...")
        threading.Thread(target=self._compare_worker, daemon=True).start()

    def _compare_worker(self) -> None:
        try:
            if self.left_temp:
                self.left_temp.cleanup()
            if self.right_temp:
                self.right_temp.cleanup()
            left_root, self.left_temp = self.resolve_source(self.left)
            right_root, self.right_temp = self.resolve_source(self.right)
            items = compare_folders(left_root, right_root, self.deep.get())
            self.after(0, lambda: self.show_items(items))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror(APP_TITLE, str(exc)))
            self.after(0, lambda: self.status.set("Compare failed"))

    def show_items(self, items: list[FolderItem]) -> None:
        self.items = items
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(items):
            self.tree.insert(
                "",
                END,
                iid=str(index),
                values=(item.rel_path, item.status, item.size_left or "", item.size_right or ""),
                tags=(item.status,),
            )
        diff_count = sum(1 for item in items if item.status != "Match")
        self.status.set(f"{len(items)} files scanned, {diff_count} differences")

    def sync(self, mode: str) -> None:
        left = self.left.path()
        right = self.right.path()
        if not left or not right or not left.is_dir() or not right.is_dir():
            messagebox.showwarning(APP_TITLE, "Folder Sync works with real folders, not compressed archives.")
            return
        if not self.items:
            messagebox.showinfo(APP_TITLE, "Run folder compare first.")
            return
        if not messagebox.askyesno(APP_TITLE, "Apply sync operation to selected differences?"):
            return
        selected = self.tree.selection() or [str(i) for i in range(len(self.items))]
        for iid in selected:
            item = self.items[int(iid)]
            src: Path | None = None
            dst: Path | None = None
            if mode in {"ltr", "mirror_ltr"} and item.left:
                src, dst = item.left, right / item.rel_path
            elif mode in {"rtl", "mirror_rtl"} and item.right:
                src, dst = item.right, left / item.rel_path
            if src and dst and src.exists():
                ensure_parent(dst)
                shutil.copy2(src, dst)
            if mode == "mirror_ltr" and item.status == "Only right" and item.right:
                item.right.unlink(missing_ok=True)
            if mode == "mirror_rtl" and item.status == "Only left" and item.left:
                item.left.unlink(missing_ok=True)
        self.start_compare()


class TextCompareTab(ttk.Frame):
    def __init__(self, parent, settings: dict):
        super().__init__(parent)
        self.settings = settings
        self.ignore_ws = BooleanVar(value=False)
        self.ignore_blank = BooleanVar(value=False)
        self.ignore_comments = BooleanVar(value=False)
        self.wrap = BooleanVar(value=True)
        self._build()

    def _build(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=X, padx=8, pady=8)
        self.left = PathPicker(top, "Left")
        self.left.pack(fill=X, pady=2)
        self.right = PathPicker(top, "Right")
        self.right.pack(fill=X, pady=2)
        controls = ttk.Frame(self)
        controls.pack(fill=X, padx=8)
        ttk.Button(controls, text="Load", command=self.load).pack(side=LEFT, padx=2)
        ttk.Button(controls, text="Compare", command=self.compare).pack(side=LEFT, padx=2)
        ttk.Button(controls, text="Save left", command=lambda: self.save("left")).pack(side=LEFT, padx=2)
        ttk.Button(controls, text="Save right", command=lambda: self.save("right")).pack(side=LEFT, padx=2)
        ttk.Checkbutton(controls, text="Ignore whitespace", variable=self.ignore_ws).pack(side=LEFT, padx=8)
        ttk.Checkbutton(controls, text="Ignore blank lines", variable=self.ignore_blank).pack(side=LEFT)
        ttk.Checkbutton(controls, text="Ignore comments", variable=self.ignore_comments).pack(side=LEFT, padx=8)
        ttk.Checkbutton(controls, text="Word wrap", variable=self.wrap, command=self.apply_wrap).pack(side=LEFT)

        panes = ttk.PanedWindow(self, orient=HORIZONTAL)
        panes.pack(fill=BOTH, expand=True, padx=8, pady=8)
        self.left_text = ScrolledText(panes, undo=True, wrap="word")
        self.right_text = ScrolledText(panes, undo=True, wrap="word")
        panes.add(self.left_text, weight=1)
        panes.add(self.right_text, weight=1)
        self.apply_settings()
        for widget in (self.left_text, self.right_text):
            widget.tag_configure("keyword", foreground="#0b5cad")
            widget.tag_configure("string", foreground="#8a4b08")
            widget.tag_configure("comment", foreground="#2f7d32")
            widget.tag_configure("number", foreground="#7b2cbf")

    def apply_settings(self) -> None:
        self.left_text.tag_configure("diff", background=self.settings["text_diff_color"])
        self.right_text.tag_configure("diff", background=self.settings["text_diff_color"])
        self.left_text.tag_configure("delete", background=self.settings["text_delete_color"])
        self.right_text.tag_configure("insert", background=self.settings["text_insert_color"])

    def apply_wrap(self) -> None:
        mode = "word" if self.wrap.get() else "none"
        self.left_text.configure(wrap=mode)
        self.right_text.configure(wrap=mode)

    def load(self) -> None:
        try:
            left = self.left.path()
            right = self.right.path()
            if left:
                self.left_text.delete("1.0", END)
                self.left_text.insert("1.0", read_text_file(left))
            if right:
                self.right_text.delete("1.0", END)
                self.right_text.insert("1.0", read_text_file(right))
            self.highlight_syntax()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def save(self, side: str) -> None:
        try:
            path = self.left.path() if side == "left" else self.right.path()
            text = self.left_text if side == "left" else self.right_text
            if not path:
                path = Path(filedialog.asksaveasfilename())
            if path:
                write_text_file(path, text.get("1.0", "end-1c"))
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def compare(self) -> None:
        for widget in (self.left_text, self.right_text):
            widget.tag_remove("diff", "1.0", END)
            widget.tag_remove("delete", "1.0", END)
            widget.tag_remove("insert", "1.0", END)
        self.highlight_syntax()
        left_raw = self.left_text.get("1.0", "end-1c")
        right_raw = self.right_text.get("1.0", "end-1c")
        left_lines = normalize_text(left_raw, self.ignore_ws.get(), self.ignore_blank.get(), self.ignore_comments.get())
        right_lines = normalize_text(right_raw, self.ignore_ws.get(), self.ignore_blank.get(), self.ignore_comments.get())
        matcher = difflib.SequenceMatcher(None, left_lines, right_lines)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                continue
            left_tag = "delete" if tag == "delete" else "diff"
            right_tag = "insert" if tag == "insert" else "diff"
            if i1 < i2:
                self.left_text.tag_add(left_tag, f"{i1 + 1}.0", f"{i2}.end")
            if j1 < j2:
                self.right_text.tag_add(right_tag, f"{j1 + 1}.0", f"{j2}.end")

    def highlight_syntax(self) -> None:
        for widget, picker in ((self.left_text, self.left), (self.right_text, self.right)):
            for tag in ("keyword", "string", "comment", "number"):
                widget.tag_remove(tag, "1.0", END)
            path = picker.path()
            suffix = path.suffix.lower() if path else ""
            text = widget.get("1.0", "end-1c")
            if suffix in {".py", ".pyw"}:
                keywords = set(keyword.kwlist)
                comment_pattern = r"#.*"
            elif suffix in {".js", ".ts", ".java", ".c", ".cpp", ".h", ".hpp", ".cs"}:
                keywords = {
                    "break",
                    "case",
                    "class",
                    "const",
                    "continue",
                    "default",
                    "else",
                    "enum",
                    "for",
                    "function",
                    "if",
                    "import",
                    "include",
                    "int",
                    "let",
                    "new",
                    "private",
                    "protected",
                    "public",
                    "return",
                    "static",
                    "struct",
                    "switch",
                    "this",
                    "var",
                    "void",
                    "while",
                }
                comment_pattern = r"//.*|/\*.*?\*/"
            elif suffix in {".html", ".htm", ".xml"}:
                keywords = {"html", "head", "body", "div", "span", "script", "style", "table", "tr", "td", "class", "id"}
                comment_pattern = r"<!--.*?-->"
            else:
                keywords = set()
                comment_pattern = r"#.*|//.*"
            self.apply_regex_tag(widget, text, r"(['\"])(?:\\.|(?!\1).)*\1", "string", re.MULTILINE)
            self.apply_regex_tag(widget, text, r"\b\d+(?:\.\d+)?\b", "number")
            if keywords:
                self.apply_regex_tag(widget, text, r"\b(" + "|".join(re.escape(word) for word in sorted(keywords)) + r")\b", "keyword")
            block_flags = re.MULTILINE | re.DOTALL if "/*" in comment_pattern or "<!--" in comment_pattern else re.MULTILINE
            self.apply_regex_tag(widget, text, comment_pattern, "comment", block_flags)

    def apply_regex_tag(self, widget: ScrolledText, text: str, pattern: str, tag: str, flags: int = re.MULTILINE) -> None:
        for match in re.finditer(pattern, text, flags=flags):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            widget.tag_add(tag, start, end)


class TableCompareTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self._build()

    def _build(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=X, padx=8, pady=8)
        self.left = PathPicker(top, "Left")
        self.left.pack(fill=X, pady=2)
        self.right = PathPicker(top, "Right")
        self.right.pack(fill=X, pady=2)
        ttk.Button(top, text="Compare table", command=self.compare).pack(anchor=W, pady=4)
        self.tree = ttk.Treeview(self, columns=("cell", "left", "right"), show="headings")
        self.tree.heading("cell", text="Cell")
        self.tree.heading("left", text="Left")
        self.tree.heading("right", text="Right")
        self.tree.column("cell", width=120)
        self.tree.pack(fill=BOTH, expand=True, padx=8, pady=8)

    def read_table(self, path: Path) -> list[list[str]]:
        if path.suffix.lower() == ".csv":
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                return [[str(cell) for cell in row] for row in csv.reader(handle)]
        if path.suffix.lower() == ".xlsx":
            if openpyxl is None:
                raise RuntimeError("Install openpyxl to compare .xlsx files: pip install openpyxl")
            workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
            sheet = workbook.active
            return [["" if cell is None else str(cell) for cell in row] for row in sheet.iter_rows(values_only=True)]
        raise RuntimeError("Supported table formats: .csv, .xlsx")

    def compare(self) -> None:
        try:
            left = self.left.path()
            right = self.right.path()
            if not left or not right:
                raise ValueError("Please choose both files.")
            left_rows = self.read_table(left)
            right_rows = self.read_table(right)
            self.tree.delete(*self.tree.get_children())
            max_rows = max(len(left_rows), len(right_rows))
            for row_index in range(max_rows):
                left_row = left_rows[row_index] if row_index < len(left_rows) else []
                right_row = right_rows[row_index] if row_index < len(right_rows) else []
                max_cols = max(len(left_row), len(right_row))
                for col_index in range(max_cols):
                    left_value = left_row[col_index] if col_index < len(left_row) else ""
                    right_value = right_row[col_index] if col_index < len(right_row) else ""
                    if left_value != right_value:
                        cell = f"R{row_index + 1}C{col_index + 1}"
                        self.tree.insert("", END, values=(cell, left_value, right_value))
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))


class PictureCompareTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.mode = StringVar(value="side")
        self.alpha = IntVar(value=50)
        self.photo = None
        self._build()

    def _build(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=X, padx=8, pady=8)
        self.left = PathPicker(top, "Left")
        self.left.pack(fill=X, pady=2)
        self.right = PathPicker(top, "Right")
        self.right.pack(fill=X, pady=2)
        controls = ttk.Frame(self)
        controls.pack(fill=X, padx=8)
        ttk.Radiobutton(controls, text="Side by side", variable=self.mode, value="side").pack(side=LEFT)
        ttk.Radiobutton(controls, text="Overlay", variable=self.mode, value="overlay").pack(side=LEFT, padx=8)
        ttk.Radiobutton(controls, text="Diff pixels", variable=self.mode, value="diff").pack(side=LEFT)
        ttk.Label(controls, text="Overlay %").pack(side=LEFT, padx=(16, 4))
        ttk.Scale(controls, from_=0, to=100, variable=self.alpha, orient=HORIZONTAL).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(controls, text="Render", command=self.render).pack(side=RIGHT, padx=4)
        self.canvas = ttk.Label(self, anchor="center")
        self.canvas.pack(fill=BOTH, expand=True, padx=8, pady=8)

    def render(self) -> None:
        if Image is None:
            messagebox.showerror(APP_TITLE, "Install Pillow to compare images: pip install pillow")
            return
        try:
            left_path = self.left.path()
            right_path = self.right.path()
            if not left_path or not right_path:
                raise ValueError("Please choose both images.")
            left_img = Image.open(left_path).convert("RGBA")
            right_img = Image.open(right_path).convert("RGBA").resize(left_img.size)
            mode = self.mode.get()
            if mode == "overlay":
                image = Image.blend(left_img, right_img, self.alpha.get() / 100)
            elif mode == "diff":
                diff = ImageChops.difference(left_img, right_img)
                image = Image.new("RGBA", left_img.size, (255, 255, 255, 255))
                image.paste((255, 0, 80, 255), mask=diff.convert("L").point(lambda p: 255 if p else 0))
            else:
                image = Image.new("RGBA", (left_img.width + right_img.width, max(left_img.height, right_img.height)), "white")
                image.paste(left_img, (0, 0))
                image.paste(right_img, (left_img.width, 0))
            image.thumbnail((1100, 720))
            self.photo = ImageTk.PhotoImage(image)
            self.canvas.configure(image=self.photo)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))


class HexCompareTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self._build()

    def _build(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=X, padx=8, pady=8)
        self.left = PathPicker(top, "Left")
        self.left.pack(fill=X, pady=2)
        self.right = PathPicker(top, "Right")
        self.right.pack(fill=X, pady=2)
        ttk.Button(top, text="Compare bytes", command=self.compare).pack(anchor=W, pady=4)
        panes = ttk.PanedWindow(self, orient=HORIZONTAL)
        panes.pack(fill=BOTH, expand=True, padx=8, pady=8)
        self.left_text = ScrolledText(panes, width=80)
        self.right_text = ScrolledText(panes, width=80)
        panes.add(self.left_text, weight=1)
        panes.add(self.right_text, weight=1)
        self.left_text.tag_configure("diff", background="#fff0cf")
        self.right_text.tag_configure("diff", background="#fff0cf")

    def format_hex(self, data: bytes) -> list[str]:
        lines = []
        for offset in range(0, len(data), 16):
            chunk = data[offset : offset + 16]
            hex_part = " ".join(f"{byte:02x}" for byte in chunk).ljust(47)
            ascii_part = "".join(chr(byte) if 32 <= byte < 127 else "." for byte in chunk)
            lines.append(f"{offset:08x}  {hex_part}  {ascii_part}")
        return lines

    def compare(self) -> None:
        try:
            left = self.left.path()
            right = self.right.path()
            if not left or not right:
                raise ValueError("Please choose both files.")
            left_data = left.read_bytes()
            right_data = right.read_bytes()
            left_lines = self.format_hex(left_data)
            right_lines = self.format_hex(right_data)
            self.left_text.delete("1.0", END)
            self.right_text.delete("1.0", END)
            self.left_text.insert("1.0", "\n".join(left_lines))
            self.right_text.insert("1.0", "\n".join(right_lines))
            for widget in (self.left_text, self.right_text):
                widget.tag_remove("diff", "1.0", END)
            max_len = max(len(left_data), len(right_data))
            for offset in range(0, max_len, 16):
                if left_data[offset : offset + 16] != right_data[offset : offset + 16]:
                    line = offset // 16 + 1
                    self.left_text.tag_add("diff", f"{line}.0", f"{line}.end")
                    self.right_text.tag_add("diff", f"{line}.0", f"{line}.end")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))


class SettingsTab(ttk.Frame):
    COLOR_FIELDS = [
        ("match_color", "Folder match"),
        ("diff_color", "Folder different"),
        ("only_left_color", "Only left"),
        ("only_right_color", "Only right"),
        ("text_diff_color", "Text changed"),
        ("text_delete_color", "Text deleted"),
        ("text_insert_color", "Text inserted"),
    ]

    def __init__(self, parent, settings: dict, on_save):
        super().__init__(parent)
        self.settings = settings
        self.on_save = on_save
        self.color_vars = {key: StringVar(value=settings[key]) for key, _ in self.COLOR_FIELDS}
        self.previews: dict[str, ttk.Label] = {}
        self.startup = BooleanVar(value=is_windows_startup_enabled())
        self._build()

    def _build(self) -> None:
        wrapper = ttk.Frame(self)
        wrapper.pack(fill=BOTH, expand=True, padx=14, pady=14)
        ttk.Label(wrapper, text="Colors").pack(anchor=W, pady=(0, 8))
        for key, label in self.COLOR_FIELDS:
            row = ttk.Frame(wrapper)
            row.pack(fill=X, pady=3)
            ttk.Label(row, text=label, width=18).pack(side=LEFT)
            preview = ttk.Label(row, width=10, textvariable=self.color_vars[key], background=self.color_vars[key].get())
            self.previews[key] = preview
            preview.pack(side=LEFT, padx=(0, 8))
            ttk.Button(row, text="Choose", command=lambda k=key, p=preview: self.choose_color(k, p)).pack(side=LEFT)

        ttk.Separator(wrapper).pack(fill=X, pady=14)
        startup_state = "normal" if os.name == "nt" else "disabled"
        ttk.Checkbutton(
            wrapper,
            text="Start PyCompare Studio with Windows",
            variable=self.startup,
            state=startup_state,
        ).pack(anchor=W)
        if os.name != "nt":
            ttk.Label(wrapper, text="Startup option is available on Windows.").pack(anchor=W, pady=(4, 0))

        buttons = ttk.Frame(wrapper)
        buttons.pack(fill=X, pady=16)
        ttk.Button(buttons, text="Save settings", command=self.save).pack(side=LEFT)
        ttk.Button(buttons, text="Reset colors", command=self.reset).pack(side=LEFT, padx=8)

    def choose_color(self, key: str, preview: ttk.Label) -> None:
        _, selected = colorchooser.askcolor(color=self.color_vars[key].get(), title="Choose color")
        if selected:
            self.color_vars[key].set(selected)
            preview.configure(background=selected)

    def save(self) -> None:
        for key in self.color_vars:
            self.settings[key] = self.color_vars[key].get()
        self.settings["startup_windows"] = self.startup.get()
        try:
            set_windows_startup(self.startup.get())
            save_settings(self.settings)
            self.on_save()
            messagebox.showinfo(APP_TITLE, "Settings saved.")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def reset(self) -> None:
        for key, _ in self.COLOR_FIELDS:
            self.color_vars[key].set(DEFAULT_SETTINGS[key])
            self.previews[key].configure(background=DEFAULT_SETTINGS[key])


class App(Tk):
    def __init__(self, left_path: str | None = None, right_path: str | None = None):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1250x820")
        self.minsize(980, 620)
        self.settings = load_settings()
        self._style()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=True)
        self.folder_tab = FolderCompareTab(self.notebook, self.settings)
        self.text_tab = TextCompareTab(self.notebook, self.settings)
        self.table_tab = TableCompareTab(self.notebook)
        self.picture_tab = PictureCompareTab(self.notebook)
        self.hex_tab = HexCompareTab(self.notebook)
        self.settings_tab = SettingsTab(self.notebook, self.settings, self.apply_settings)
        self.notebook.add(self.folder_tab, text="Folder Compare & Sync")
        self.notebook.add(self.text_tab, text="Text Compare")
        self.notebook.add(self.table_tab, text="Table Compare")
        self.notebook.add(self.picture_tab, text="Picture Compare")
        self.notebook.add(self.hex_tab, text="Hex Compare")
        self.notebook.add(self.settings_tab, text="Settings")
        if left_path and right_path:
            self.after(150, lambda: self.open_compare(left_path, right_path))

    def _style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Treeview", rowheight=24)

    def apply_settings(self) -> None:
        self.folder_tab.apply_settings()
        self.text_tab.apply_settings()

    def open_compare(self, left_path: str, right_path: str) -> None:
        left = Path(left_path)
        right = Path(right_path)
        if left.is_dir() or right.is_dir():
            self.notebook.select(self.folder_tab)
            self.folder_tab.left.value.set(str(left))
            self.folder_tab.right.value.set(str(right))
            self.folder_tab.start_compare()
            return

        left_suffix = left.suffix.lower()
        right_suffix = right.suffix.lower()
        if left_suffix in TABLE_EXTENSIONS and right_suffix in TABLE_EXTENSIONS:
            self.notebook.select(self.table_tab)
            self.table_tab.left.value.set(str(left))
            self.table_tab.right.value.set(str(right))
            self.table_tab.compare()
        elif left_suffix in PICTURE_EXTENSIONS and right_suffix in PICTURE_EXTENSIONS:
            self.notebook.select(self.picture_tab)
            self.picture_tab.left.value.set(str(left))
            self.picture_tab.right.value.set(str(right))
            self.picture_tab.render()
        elif self.looks_like_text(left) and self.looks_like_text(right):
            self.notebook.select(self.text_tab)
            self.text_tab.left.value.set(str(left))
            self.text_tab.right.value.set(str(right))
            self.text_tab.load()
            self.text_tab.compare()
        else:
            self.notebook.select(self.hex_tab)
            self.hex_tab.left.value.set(str(left))
            self.hex_tab.right.value.set(str(right))
            self.hex_tab.compare()

    def looks_like_text(self, path: Path) -> bool:
        if path.suffix.lower() in TEXT_EXTENSIONS:
            return True
        mime_type, _ = mimetypes.guess_type(path)
        if mime_type and mime_type.startswith("text/"):
            return True
        try:
            sample = path.read_bytes()[:4096]
        except OSError:
            return False
        return b"\x00" not in sample


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PyCompare Studio")
    parser.add_argument("--left", help="Left file or folder to compare.")
    parser.add_argument("--right", help="Right file or folder to compare.")
    parser.add_argument("--select-left", help="Save this path as the left side for shell compare.")
    parser.add_argument("--compare-to-left", help="Compare this path with the saved left path.")
    parser.add_argument("--compare-selected", nargs="+", help="Compare exactly two selected paths.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if run_context_action(args):
        return
    app = App(args.left, args.right)
    app.mainloop()


if __name__ == "__main__":
    main()
