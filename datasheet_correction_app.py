import copy
import os
import re
import sys
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET


APP_TITLE = "Datasheet Correction"
APP_VERSION = "v0.1"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)


def qn(tag):
    return f"{{{W_NS}}}{tag}"


FIELD_GROUPS = {
    "Basic": [
        "Customer",
        "Project Name",
        "Contact",
        "Service",
        "Item No.",
        "Datasheet No.",
        "Date",
        "Designed by",
    ],
    "Duty Requirements": [
        "Media",
        "Heat capacity",
        "Temperature inlet",
        "Temperature outlet",
        "Mass flow rate",
        "Volume flowrate",
        "Operating Pressure",
        "Pressure drop",
    ],
    "Unit Data": [
        "Design Temperature",
        "Design / Test Pressure",
        "Connection Type",
        "O.H.T.C",
        "Heat transfer area",
        "Number of plates",
        "Surface margin",
    ],
    "Product Properties": [
        "Density (Liquid)",
        "Specific heat capacity (Liquid)",
        "Thermal conductivity (Liquid)",
        "Viscosity (Liquid)",
        "Latent heat",
    ],
}

FIELD_CHOICES = [f"{group} > {name}" for group, names in FIELD_GROUPS.items() for name in names]

UNIT_OPTIONS = {
    "Heat capacity": ["kcal/h -> kW"],
    "O.H.T.C": ["kcal/(m²·h·K) -> W/(m²·K)"],
    "Specific heat capacity (Liquid)": ["kcal/(kg·K) -> kJ/(kg·K)"],
    "Thermal conductivity (Liquid)": ["kcal/(m·h·K) -> W/(m·K)"],
    "Latent heat": ["kcal/kg -> kJ/kg"],
    "Operating Pressure": ["bar_G -> bar_A", "bar_G -> kPaG", "bar_G -> kPaA"],
    "Pressure drop": ["kPa -> bar", "kPa -> bar_G"],
    "Design / Test Pressure": ["bar -> kPa", "bar -> bar_A"],
}


def normalize(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def clean_label(text):
    return normalize(text).replace(" :", "").replace(":", "").lower()


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def default_scan_dir():
    base = Path(app_dir())
    if getattr(sys, "frozen", False) and base.name.lower() == "dist":
        return str(base.parent)
    return str(base)


class DatasheetFile:
    def __init__(self, source_path, display_name):
        self.source_path = source_path
        self.display_name = display_name
        self.kind = ""
        self.form = ""
        self.result = ""
        self.selected = True

    @property
    def key(self):
        return self.source_path


class DocxPackage:
    def __init__(self, data):
        self.data = data
        self.entries = {}
        with zipfile.ZipFile(BytesIO(data), "r") as zf:
            for info in zf.infolist():
                self.entries[info.filename] = zf.read(info.filename)

    def xml_parts(self):
        names = []
        for name in self.entries:
            if re.match(r"word/(document|header|footer|footnotes|endnotes).*\.xml$", name):
                names.append(name)
        return names

    def read_xml(self, name):
        return ET.fromstring(self.entries[name])

    def write_xml(self, name, root):
        self.entries[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def to_bytes(self):
        out = BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in self.entries.items():
                zf.writestr(name, data)
        return out.getvalue()


def load_docx_bytes(item):
    with open(item.source_path, "rb") as f:
        return f.read()


def iter_rows(root):
    for row in root.iter(qn("tr")):
        cells = list(row.findall(qn("tc")))
        yield row, cells


def cell_text(cell):
    return "".join(node.text or "" for node in cell.iter(qn("t")))


def row_texts(row):
    return [normalize(cell_text(cell)) for cell in row.findall(qn("tc"))]


def set_cell_text(cell, value):
    texts = list(cell.iter(qn("t")))
    if not texts:
        para = cell.find(qn("p"))
        if para is None:
            para = ET.SubElement(cell, qn("p"))
        run = ET.SubElement(para, qn("r"))
        text = ET.SubElement(run, qn("t"))
        text.text = str(value)
        return
    texts[0].text = str(value)
    for node in texts[1:]:
        node.text = ""


def replace_cell_text(cell, pattern, repl):
    old = cell_text(cell)
    new = re.sub(pattern, repl, old)
    if new != old:
        set_cell_text(cell, new)
        return old, new
    return None


def row_label_matches(text, target):
    label = clean_label(text)
    target_label = clean_label(target)
    return label == target_label or target_label in label


def first_number_format(old_text, new_value):
    value = float(new_value)
    match = re.search(r"-?\d+(?:\.\d+)?", str(old_text))
    if not match:
        return str(new_value)
    original = match.group(0)
    if "." in original:
        decimals = len(original.split(".")[1])
        return f"{value:.{decimals}f}"
    return str(int(round(value)))


def convert_value(value_text, rule):
    values = re.findall(r"-?\d+(?:\.\d+)?", value_text)
    if not values:
        return value_text

    def repl(match):
        number = float(match.group(0))
        if rule == "kcal/h -> kW":
            converted = number * 0.001163
        elif rule in ("kcal/(m²·h·K) -> W/(m²·K)", "kcal/(m·h·K) -> W/(m·K)"):
            converted = number * 1.163
        elif rule in ("kcal/(kg·K) -> kJ/(kg·K)", "kcal/kg -> kJ/kg"):
            converted = number * 4.1868
        elif rule == "bar_G -> bar_A":
            converted = number + 1.01325
        elif rule == "bar_G -> kPaG":
            converted = number * 100.0
        elif rule == "bar_G -> kPaA":
            converted = (number + 1.01325) * 100.0
        elif rule in ("kPa -> bar", "kPa -> bar_G"):
            converted = number / 100.0
        elif rule in ("bar -> kPa", "bar -> bar_A"):
            converted = number * 100.0 if rule == "bar -> kPa" else number + 1.01325
        else:
            converted = number
        decimals = 2
        if abs(converted) >= 100:
            decimals = 1
        if abs(converted) < 10:
            decimals = 2
        return f"{converted:.{decimals}f}"

    return re.sub(r"-?\d+(?:\.\d+)?", repl, value_text)


def target_unit(rule):
    return rule.split(" -> ", 1)[1]


class DatasheetEditor:
    def __init__(self, item):
        self.item = item
        self.package = DocxPackage(load_docx_bytes(item))
        self.changes = []

    def scan_summary(self):
        text = self.all_text()
        result = re.search(r"MG PHE Result\s*:\s*([^\n\r]+?)(?:Series|Section|1\.)", text)
        return normalize(result.group(1)) if result else ""

    def all_text(self):
        chunks = []
        for part in self.package.xml_parts():
            root = self.package.read_xml(part)
            chunks.append(" ".join(node.text or "" for node in root.iter(qn("t"))))
        return normalize(" ".join(chunks))

    def apply_task(self, task):
        before = len(self.changes)
        for part in self.package.xml_parts():
            root = self.package.read_xml(part)
            changed = self.apply_task_to_root(root, task)
            if changed:
                self.package.write_xml(part, root)
        return self.changes[before:]

    def apply_task_to_root(self, root, task):
        field = task["field"]
        mode = task.get("mode", "set")
        if field in {"Customer", "Project Name", "Contact", "Service", "Item No.", "Date"}:
            return self.set_header_pair(root, field, task["value"])
        if field == "Datasheet No.":
            return self.replace_in_cells(root, r"(Datasheet No\.\s*:\s*).*", rf"\g<1>{task['value']}", field)
        if field == "Designed by":
            return self.replace_in_text_nodes(root, r"(-?\s*Designed by\s+).*", rf"\g<1>{task['value']}", field)
        if field in {"Media", "Temperature inlet", "Temperature outlet", "Mass flow rate", "Volume flowrate", "Operating Pressure", "Pressure drop"}:
            if mode == "convert":
                return self.convert_row(root, field, task["unit_rule"], two_sided=True)
            return self.set_two_sided_row(root, field, task.get("hot", ""), task.get("cold", ""), task.get("unit", ""))
        if field == "Connection Type":
            return self.set_two_sided_row(root, field, task.get("hot", ""), task.get("cold", ""), "")
        if field in {"Design Temperature", "Design / Test Pressure", "Heat transfer area", "Number of plates", "Surface margin"}:
            label = "Design temperature" if field == "Design Temperature" else field
            if mode == "convert":
                return self.convert_row(root, label, task["unit_rule"], two_sided=False)
            return self.set_single_row(root, label, task.get("value", ""), task.get("unit", ""))
        if field in {"Heat capacity", "O.H.T.C", "Density (Liquid)", "Specific heat capacity (Liquid)", "Thermal conductivity (Liquid)", "Viscosity (Liquid)", "Latent heat"}:
            if mode == "convert":
                return self.convert_row(root, field, task["unit_rule"], two_sided=("capacity" not in field.lower() or field == "Heat capacity"))
            return self.set_single_or_two(root, field, task)
        return False

    def set_header_pair(self, root, field, value):
        changed = False
        for _row, cells in iter_rows(root):
            texts = [cell_text(c) for c in cells]
            for idx, text in enumerate(texts[:-1]):
                if row_label_matches(text, field):
                    old = cell_text(cells[idx + 1])
                    if old != value:
                        set_cell_text(cells[idx + 1], value)
                        self.changes.append((field, old, value))
                        changed = True
        return changed

    def replace_in_cells(self, root, pattern, repl, field):
        changed = False
        for cell in root.iter(qn("tc")):
            result = replace_cell_text(cell, pattern, repl)
            if result:
                old, new = result
                self.changes.append((field, old, new))
                changed = True
        return changed

    def replace_in_text_nodes(self, root, pattern, repl, field):
        changed = False
        for node in root.iter(qn("t")):
            old = node.text or ""
            new = re.sub(pattern, repl, old)
            if old != new:
                node.text = new
                self.changes.append((field, old, new))
                changed = True
        return changed

    def find_rows(self, root, label):
        for _row, cells in iter_rows(root):
            if not cells:
                continue
            if row_label_matches(cell_text(cells[0]), label):
                yield cells

    def set_two_sided_row(self, root, label, hot, cold, unit):
        changed = False
        for cells in self.find_rows(root, label):
            if len(cells) > 1 and hot:
                old = cell_text(cells[1])
                if old != hot:
                    set_cell_text(cells[1], hot)
                    self.changes.append((label + " Hot", old, hot))
                    changed = True
            if len(cells) > 2 and cold:
                old = cell_text(cells[2])
                if old != cold:
                    set_cell_text(cells[2], cold)
                    self.changes.append((label + " Cold", old, cold))
                    changed = True
            if unit and len(cells) > 3:
                old = cell_text(cells[3])
                if old != unit:
                    set_cell_text(cells[3], unit)
                    self.changes.append((label + " Unit", old, unit))
                    changed = True
        return changed

    def set_single_row(self, root, label, value, unit):
        changed = False
        for cells in self.find_rows(root, label):
            if len(cells) > 1 and value:
                old = cell_text(cells[1])
                if old != value:
                    set_cell_text(cells[1], value)
                    self.changes.append((label, old, value))
                    changed = True
            if unit and len(cells) > 2:
                old = cell_text(cells[-1])
                if old != unit:
                    set_cell_text(cells[-1], unit)
                    self.changes.append((label + " Unit", old, unit))
                    changed = True
        return changed

    def set_single_or_two(self, root, label, task):
        if task.get("hot") or task.get("cold"):
            return self.set_two_sided_row(root, label, task.get("hot", ""), task.get("cold", ""), task.get("unit", ""))
        return self.set_single_row(root, label, task.get("value", ""), task.get("unit", ""))

    def convert_row(self, root, label, rule, two_sided=True):
        changed = False
        for cells in self.find_rows(root, label):
            value_indexes = [1, 2] if two_sided and len(cells) > 3 else [1]
            for idx in value_indexes:
                if idx >= len(cells):
                    continue
                old = cell_text(cells[idx])
                new = convert_value(old, rule)
                if old != new:
                    set_cell_text(cells[idx], new)
                    self.changes.append((label, old, new))
                    changed = True
            unit_idx = len(cells) - 1
            if unit_idx > 0:
                old_unit = cell_text(cells[unit_idx])
                new_unit = target_unit(rule)
                if old_unit != new_unit:
                    set_cell_text(cells[unit_idx], new_unit)
                    self.changes.append((label + " Unit", old_unit, new_unit))
                    changed = True
        return changed

    def save_docx(self, path):
        with open(path, "wb") as f:
            f.write(self.package.to_bytes())

    def to_bytes(self):
        return self.package.to_bytes()


def discover_files(folder):
    files = []
    root = Path(folder)
    skip_dirs = {".git", ".agents", ".codex", "__pycache__", "build", "dist", "output"}
    paths = []
    for path in root.rglob("*"):
        if any(part in skip_dirs for part in path.parts):
            continue
        if path.is_file() and path.suffix.lower() == ".docx":
            paths.append(path)
    for path in sorted(paths):
        if path.name.startswith("~$"):
            continue
        item = DatasheetFile(str(path), str(path.relative_to(root)))
        fill_file_info(item)
        files.append(item)
    return files


def fill_file_info(item):
    name = item.display_name
    lower = name.lower()
    item.kind = "GPHE" if "gphe" in lower else "BPHE" if "bphe" in lower else ""
    item.form = "simple" if "simple" in lower else "standard" if "standard" in lower else "internal" if "internal" in lower else ""
    try:
        editor = DatasheetEditor(item)
        item.result = editor.scan_summary()
        if not item.result:
            item.result = "summary not found"
        if not item.kind:
            if "M70" in item.result:
                item.kind = "GPHE"
            elif "MC" in item.result:
                item.kind = "BPHE"
    except zipfile.BadZipFile:
        item.result = "not a valid docx"
    except PermissionError:
        item.result = "permission denied"
    except Exception as exc:
        item.result = f"read error: {exc.__class__.__name__}"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1180x760")
        self.minsize(1050, 680)
        self.files = []
        self.tasks = []
        self.vars = {}
        self.build_ui()

    def build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=8)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        ttk.Button(top, text="Folder", command=self.choose_folder).grid(row=0, column=0, padx=(0, 6))
        self.folder_var = tk.StringVar(value=default_scan_dir())
        ttk.Entry(top, textvariable=self.folder_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="Load", command=self.load_folder).grid(row=0, column=2, padx=(6, 0))
        self.status_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.status_var).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        main = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

        left = ttk.Frame(main, padding=6)
        main.add(left, weight=3)
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)
        button_bar = ttk.Frame(left)
        button_bar.grid(row=0, column=0, sticky="ew")
        ttk.Button(button_bar, text="All", command=lambda: self.set_all_files(True)).pack(side=tk.LEFT)
        ttk.Button(button_bar, text="None", command=lambda: self.set_all_files(False)).pack(side=tk.LEFT, padx=4)
        columns = ("use", "name", "kind", "form", "result")
        self.file_tree = ttk.Treeview(left, columns=columns, show="headings", selectmode="browse")
        for col, width in [("use", 55), ("name", 300), ("kind", 70), ("form", 80), ("result", 190)]:
            self.file_tree.heading(col, text=col)
            self.file_tree.column(col, width=width, anchor=tk.W)
        self.file_tree.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.file_tree.bind("<Double-1>", self.toggle_file)
        self.file_tree.bind("<ButtonRelease-1>", self.toggle_file_if_use_column)

        right = ttk.Frame(main, padding=6)
        main.add(right, weight=2)
        right.columnconfigure(0, weight=1)

        ttk.Label(right, text="Edit field").grid(row=0, column=0, sticky="w")
        self.field_var = tk.StringVar(value=FIELD_CHOICES[0])
        self.field_combo = ttk.Combobox(right, textvariable=self.field_var, values=FIELD_CHOICES, state="readonly")
        self.field_combo.grid(row=1, column=0, sticky="ew", pady=(2, 8))
        self.field_combo.bind("<<ComboboxSelected>>", lambda _e: self.render_field_form())

        self.form = ttk.LabelFrame(right, text="Input", padding=8)
        self.form.grid(row=2, column=0, sticky="ew")
        self.form.columnconfigure(1, weight=1)

        task_buttons = ttk.Frame(right)
        task_buttons.grid(row=3, column=0, sticky="ew", pady=8)
        ttk.Button(task_buttons, text="Add Task", command=self.add_task).pack(side=tk.LEFT)
        ttk.Button(task_buttons, text="Remove Task", command=self.remove_task).pack(side=tk.LEFT, padx=4)
        ttk.Button(task_buttons, text="Preview", command=self.preview).pack(side=tk.LEFT, padx=4)
        ttk.Button(task_buttons, text="Apply", command=self.apply).pack(side=tk.RIGHT)

        ttk.Label(right, text="Task list").grid(row=4, column=0, sticky="w")
        self.task_list = tk.Listbox(right, height=9)
        self.task_list.grid(row=5, column=0, sticky="nsew")
        right.rowconfigure(5, weight=1)

        self.preview_text = tk.Text(self, height=10, wrap=tk.NONE)
        self.preview_text.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.render_field_form()
        self.load_folder()

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get() or default_scan_dir())
        if folder:
            self.folder_var.set(folder)
            self.load_folder()

    def load_folder(self):
        folder = self.folder_var.get()
        try:
            self.files = discover_files(folder) if os.path.isdir(folder) else []
        except Exception as exc:
            self.files = []
            self.status_var.set(f"Load failed: {exc}")
            messagebox.showerror(APP_TITLE, f"Failed to load folder:\n{folder}\n\n{exc}")
            self.refresh_file_tree()
            return
        self.refresh_file_tree()
        self.status_var.set(f"Loaded {len(self.files)} datasheet file(s) from {folder}")

    def refresh_file_tree(self):
        self.file_tree.delete(*self.file_tree.get_children())
        for idx, item in enumerate(self.files):
            use = "☑" if item.selected else "☐"
            self.file_tree.insert("", "end", iid=str(idx), values=(use, item.display_name, item.kind, item.form, item.result))

    def toggle_file_if_use_column(self, event):
        region = self.file_tree.identify("region", event.x, event.y)
        col = self.file_tree.identify_column(event.x)
        if region == "cell" and col == "#1":
            self.toggle_file()

    def toggle_file(self, _event=None):
        iid = self.file_tree.focus()
        if iid == "":
            return
        item = self.files[int(iid)]
        item.selected = not item.selected
        self.refresh_file_tree()

    def set_all_files(self, selected):
        for item in self.files:
            item.selected = selected
        self.refresh_file_tree()

    def selected_field(self):
        return self.field_var.get().split(" > ", 1)[1]

    def clear_form(self):
        for child in self.form.winfo_children():
            child.destroy()
        self.vars = {}

    def add_entry(self, row, label, key, default=""):
        ttk.Label(self.form, text=label).grid(row=row, column=0, sticky="w", pady=2)
        var = tk.StringVar(value=default)
        ttk.Entry(self.form, textvariable=var).grid(row=row, column=1, sticky="ew", pady=2)
        self.vars[key] = var

    def add_combo(self, row, label, key, values):
        ttk.Label(self.form, text=label).grid(row=row, column=0, sticky="w", pady=2)
        var = tk.StringVar(value=values[0] if values else "")
        ttk.Combobox(self.form, textvariable=var, values=values, state="readonly").grid(row=row, column=1, sticky="ew", pady=2)
        self.vars[key] = var

    def render_field_form(self):
        self.clear_form()
        field = self.selected_field()
        unit_rules = UNIT_OPTIONS.get(field, [])
        row = 0
        if unit_rules:
            self.add_combo(row, "Mode", "mode", ["Set value", "Convert unit"])
            self.vars["mode"].trace_add("write", lambda *_: self.render_field_form_mode())
            row += 1
        self.render_field_form_mode(row)

    def render_field_form_mode(self, start_row=1):
        field = self.selected_field()
        mode = self.vars.get("mode").get() if self.vars.get("mode") else "Set value"
        for child in list(self.form.winfo_children()):
            info = child.grid_info()
            if info and int(info.get("row", 0)) >= start_row:
                child.destroy()
        keep = {k: v for k, v in self.vars.items() if k == "mode"}
        self.vars = keep
        row = start_row
        if mode == "Convert unit":
            self.add_combo(row, "Convert", "unit_rule", UNIT_OPTIONS.get(field, []))
            return
        if field in {"Media", "Connection Type", "Temperature inlet", "Temperature outlet", "Mass flow rate", "Volume flowrate", "Operating Pressure", "Pressure drop"}:
            self.add_entry(row, "Hot Side", "hot"); row += 1
            self.add_entry(row, "Cold Side", "cold"); row += 1
            if field not in {"Media", "Connection Type"}:
                self.add_entry(row, "Unit", "unit")
        elif field == "Design / Test Pressure":
            self.add_entry(row, "Design / Test", "value", "6.0 / 7.5"); row += 1
            self.add_entry(row, "Unit", "unit", "bar")
        elif field == "Design Temperature":
            self.add_entry(row, "Value", "value", "130"); row += 1
            self.add_entry(row, "Unit", "unit", "°C")
        elif field in {"Customer", "Project Name", "Contact", "Service", "Item No.", "Datasheet No.", "Date", "Designed by"}:
            default = datetime.today().strftime("%Y-%m-%d") if field == "Date" else ""
            self.add_entry(row, "Value", "value", default)
        else:
            self.add_entry(row, "Value", "value"); row += 1
            self.add_entry(row, "Unit", "unit")

    def add_task(self):
        field = self.selected_field()
        task = {"field": field}
        mode = self.vars.get("mode").get() if self.vars.get("mode") else "Set value"
        if mode == "Convert unit":
            task["mode"] = "convert"
            task["unit_rule"] = self.vars["unit_rule"].get()
        else:
            task["mode"] = "set"
            for key, var in self.vars.items():
                if key != "mode":
                    task[key] = var.get()
        self.tasks.append(task)
        self.refresh_task_list()

    def task_label(self, task):
        if task.get("mode") == "convert":
            return f"{task['field']}: {task['unit_rule']}"
        parts = [task["field"]]
        for key in ("value", "hot", "cold", "unit"):
            if task.get(key):
                parts.append(f"{key}={task[key]}")
        return " | ".join(parts)

    def refresh_task_list(self):
        self.task_list.delete(0, tk.END)
        for task in self.tasks:
            self.task_list.insert(tk.END, self.task_label(task))

    def remove_task(self):
        sel = self.task_list.curselection()
        if not sel:
            return
        del self.tasks[sel[0]]
        self.refresh_task_list()

    def preview(self):
        self.preview_text.delete("1.0", tk.END)
        if not self.tasks:
            self.preview_text.insert(tk.END, "No tasks.\n")
            return
        for item in [f for f in self.files if f.selected]:
            self.preview_text.insert(tk.END, f"== {item.display_name} ==\n")
            try:
                editor = DatasheetEditor(item)
                for task in self.tasks:
                    changes = editor.apply_task(copy.deepcopy(task))
                    for field, old, new in changes:
                        self.preview_text.insert(tk.END, f"  {field}: {old} -> {new}\n")
                if not editor.changes:
                    self.preview_text.insert(tk.END, "  no matched fields\n")
            except Exception as exc:
                self.preview_text.insert(tk.END, f"  ERROR: {exc}\n")
            self.preview_text.insert(tk.END, "\n")

    def apply(self):
        selected = [f for f in self.files if f.selected]
        if not selected or not self.tasks:
            messagebox.showwarning(APP_TITLE, "Select files and add at least one task.")
            return
        folder = self.folder_var.get()
        out_dir = os.path.join(folder, "output")
        os.makedirs(out_dir, exist_ok=True)
        count = 0
        errors = []
        for item in selected:
            try:
                editor = DatasheetEditor(item)
                for task in self.tasks:
                    editor.apply_task(copy.deepcopy(task))
                base = Path(item.source_path).stem
                out_path = os.path.join(out_dir, f"{base}_corrected.docx")
                editor.save_docx(out_path)
                count += 1
            except Exception as exc:
                errors.append(f"{item.display_name}: {exc}")
        self.preview()
        if errors:
            messagebox.showerror(APP_TITLE, f"Saved {count} file(s).\n\n" + "\n".join(errors[:8]))
        else:
            messagebox.showinfo(APP_TITLE, f"Saved {count} file(s) to:\n{out_dir}")


if __name__ == "__main__":
    App().mainloop()
