import sys
import os
import json
import copy
import tempfile
import pandas as pd
from PyQt6.QtCore import QObject, pyqtSlot, pyqtProperty, QUrl, pyqtSignal, QThread, Qt, QSettings, QTimer, QStandardPaths
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox
from PyQt6.QtQml import QQmlApplicationEngine
from PyQt6.QtGui import QIcon, QFontDatabase, QFont

os.environ.setdefault("QT_QUICK_CONTROLS_STYLE", "Basic")
os.environ.setdefault("QT_QUICK_CONTROLS_FALLBACK_STYLE", "Basic")
os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.fonts.warning=false")

def _app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _resource_base_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return getattr(sys, "_MEIPASS")
    return _app_base_dir()

APP_BASE_DIR = _app_base_dir()
RESOURCE_BASE_DIR = _resource_base_dir()

def _user_data_dir():
    candidates = []
    docs = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
    if docs:
        candidates.append(os.path.join(docs, "CubeFlow"))
    candidates.append(os.path.join(os.path.expanduser("~"), "Documents", "CubeFlow"))
    candidates.append(os.path.join(tempfile.gettempdir(), "CubeFlow"))

    for candidate in candidates:
        try:
            if not candidate:
                continue
            os.makedirs(candidate, exist_ok=True)
            probe_path = os.path.join(candidate, ".write_probe.tmp")
            with open(probe_path, "w", encoding="utf-8") as fp:
                fp.write("ok")
            os.remove(probe_path)
            return candidate
        except Exception:
            continue

    return os.path.join(tempfile.gettempdir(), "CubeFlow")

def _initial_formats_dir():
    return os.path.join(_user_data_dir(), "formats")


def ensure_xlsx_extension(file_name):
    if not file_name:
        return "output.xlsx"
    return file_name if file_name.lower().endswith(".xlsx") else f"{file_name}.xlsx"

def normalize_batch_output_name(file_name):
    raw_name = (file_name or "").strip()
    if not raw_name:
        return ""
    if raw_name.lower().endswith(".xlsx"):
        raw_name = raw_name[:-5]
    elif raw_name.lower().endswith(".xls"):
        raw_name = raw_name[:-4]
    raw_name = raw_name.strip()
    if not raw_name:
        return ""
    return f"{raw_name}.xlsx"

def get_invalid_batch_name_message(file_name):
    raw_name = (file_name or "").strip()
    if raw_name.lower().endswith(".xlsx"):
        raw_name = raw_name[:-5]
    elif raw_name.lower().endswith(".xls"):
        raw_name = raw_name[:-4]
    base_name = raw_name.strip()
    if not base_name:
        return "File name cannot be empty."
    invalid_chars = '<>:"/\\|?*'
    bad = sorted(set(ch for ch in base_name if ch in invalid_chars or ord(ch) < 32))
    if bad:
        return (
            f"Invalid character(s): {' '.join(bad)}. "
            "Use letters, numbers, spaces, '-', '_', '(', ')', '.'.\n"
            "Not allowed: < > : \" / \\ | ? *"
        )
    if base_name.endswith(".") or base_name.endswith(" "):
        return "File name cannot end with a dot or space."
    return ""

def get_invalid_output_directory_message(directory):
    folder = (directory or "").strip()
    if not folder:
        return "Save folder cannot be empty."
    if not os.path.exists(folder):
        return "Save folder does not exist."
    if not os.path.isdir(folder):
        return "Save path is not a folder."
    if not os.access(folder, os.W_OK):
        return "Save folder is not writable."
    return ""


def excel_col_to_index(col_name):
    name = (col_name or "").strip().upper()
    if not name or not name.isalpha():
        return 0
    idx = 0
    for ch in name:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def index_to_excel_col(idx):
    n = max(0, int(idx))
    out = ""
    while True:
        n, rem = divmod(n, 26)
        out = chr(ord('A') + rem) + out
        if n == 0:
            break
        n -= 1
    return out


def build_default_batch_output_path(xml_file):
    source_dir = os.path.dirname(xml_file)
    source_base = os.path.splitext(os.path.basename(xml_file))[0]
    return os.path.join(source_dir, f"{source_base}.xlsx")


def normalize_path(path):
    if not path:
        return ""
    normalized = os.path.normpath(str(path))
    return normalized


def collect_xml_files_from_paths(paths, should_stop=None):
    collected = []
    seen = set()
    for raw_path in paths:
        if callable(should_stop) and should_stop():
            return collected
        path = normalize_path(raw_path)
        if not path:
            continue
        if os.path.isfile(path):
            if path.lower().endswith(".xml") and path not in seen:
                collected.append(path)
                seen.add(path)
            continue
        if os.path.isdir(path):
            for root, _, files in os.walk(path):
                if callable(should_stop) and should_stop():
                    return collected
                for file_name in files:
                    if callable(should_stop) and should_stop():
                        return collected
                    if not file_name.lower().endswith(".xml"):
                        continue
                    full_path = os.path.join(root, file_name)
                    if full_path in seen:
                        continue
                    collected.append(full_path)
                    seen.add(full_path)
    return collected


def confirm_overwrite_paths(paths, title="Confirm Overwrite"):
    existing = [p for p in paths if os.path.exists(p)]
    if not existing:
        return True
    preview = "\n".join(existing[:8])
    suffix = "" if len(existing) <= 8 else f"\n...and {len(existing) - 8} more file(s)"
    msg = (
        "The following file(s) already exist:\n\n"
        f"{preview}{suffix}\n\n"
        "Do you want to overwrite them?"
    )
    result = QMessageBox.question(
        None,
        title,
        msg,
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No
    )
    return result == QMessageBox.StandardButton.Yes


def export_dataframe_to_excel(df, xml_type, save_path, xml_file):
    df = df.replace([float('inf'), float('-inf')], 0).fillna(0)
    is_builtin = xml_type in ["Den", "Glacier", "Globe"]
    with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
        xml_file_name = os.path.splitext(os.path.basename(xml_file))[0]
        sheet_name = ('_'.join(xml_file_name.split('_')[:-1]) if '_' in xml_file_name else xml_file_name)[:31]
        custom_start_row = int(df.attrs.get("data_start_row", 3))
        start_row = 0 if is_builtin else max(0, custom_start_row - 1)
        df.to_excel(writer, index=False, header=False, sheet_name=sheet_name, startrow=start_row)
        workbook = writer.book
        ws = writer.sheets[sheet_name]

        if not is_builtin:
            header_fmt = workbook.add_format({'num_format': '@', 'bg_color': '#99CC00', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1})
            formula_header_fmt = workbook.add_format({'num_format': '@', 'bg_color': '#B4C6E7', 'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1})
            generic_fmt = workbook.add_format({'num_format': 'General', 'border': 1, 'align': 'right'})
            formula_fmt = workbook.add_format({'num_format': '0.00', 'bg_color': '#B4C6E7', 'border': 1, 'align': 'right'})
            generic_highlight_fmt = workbook.add_format({'num_format': 'General', 'bg_color': '#FFFF00', 'border': 1, 'align': 'right'})
            formula_highlight_fmt = workbook.add_format({'num_format': '0.00', 'bg_color': '#FFFF00', 'border': 1, 'align': 'right'})
            formula_columns = df.attrs.get("formula_columns", [])
            formula_column_set = set()
            if isinstance(formula_columns, list):
                for col_idx in formula_columns:
                    try:
                        safe_idx = int(col_idx)
                    except Exception:
                        continue
                    if 0 <= safe_idx < df.shape[1]:
                        formula_column_set.add(safe_idx)

            header_row_1 = df.attrs.get("custom_header_row_1", [])
            header_row_2 = df.attrs.get("custom_header_row_2", [])
            if not isinstance(header_row_1, list):
                header_row_1 = []
            if not isinstance(header_row_2, list):
                header_row_2 = []
            for c in range(df.shape[1]):
                fmt = formula_header_fmt if c in formula_column_set else header_fmt
                text_1 = str(header_row_1[c]) if c < len(header_row_1) and header_row_1[c] is not None else ""
                text_2 = str(header_row_2[c]) if c < len(header_row_2) and header_row_2[c] is not None else ""
                ws.write(0, c, text_1, fmt)
                ws.write(1, c, text_2, fmt)

            for r in range(len(df)):
                excel_r = start_row + r
                row_highlight = False
                cell_value = df.iloc[r, 0] if df.shape[1] > 0 else None
                if isinstance(cell_value, str):
                    try:
                        hh, mm, ss = map(int, cell_value.strip().split(' ')[1].split(':'))
                        row_highlight = ss != 0
                    except Exception:
                        row_highlight = False
                for c in range(df.shape[1]):
                    val = df.iloc[r, c]
                    is_formula = isinstance(val, str) and val.startswith("=")
                    if is_formula:
                        cell_fmt = formula_highlight_fmt if row_highlight else formula_fmt
                        ws.write_formula(excel_r, c, val, cell_fmt)
                    else:
                        cell_fmt = generic_highlight_fmt if row_highlight else generic_fmt
                        ws.write(excel_r, c, val, cell_fmt)
            custom_widths = df.attrs.get("custom_widths", None)
            if isinstance(custom_widths, list):
                for i, w in enumerate(custom_widths[:df.shape[1]]):
                    ws.set_column(i, i, float(w))
            return

        general_fmt = workbook.add_format({'num_format': 'General', 'border': 1, 'align': 'right'})
        text_fmt = workbook.add_format({'num_format': '@', 'border': 1, 'align': 'right'})
        num_fmt = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'right'})
        header_fmt = workbook.add_format({'num_format': '@', 'bg_color': '#99CC00', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1})
        colored_num_fmt = workbook.add_format({'num_format': '0.00', 'bg_color': '#F2E6FF', 'border': 1, 'align': 'right'})
        black_header_fmt = workbook.add_format({'num_format': '@', 'bg_color': '#F2E6FF', 'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1})

        for r in range(2):
            for c in range(df.shape[1]):
                val = df.iloc[r, c]
                if xml_type == "Globe" and c in [3, 5, 10]:
                    ws.write(r, c, val, black_header_fmt)
                elif xml_type in ["Glacier", "Den"] and c in [3, 5, 8]:
                    ws.write(r, c, val, black_header_fmt)
                else:
                    ws.write(r, c, val, header_fmt)

        for r in range(2, len(df)):
            for c in range(df.shape[1]):
                val = df.iloc[r, c]
                if xml_type == "Den":
                    if c == 1:
                        ws.write(r, c, val, text_fmt)
                    elif c in [3, 5, 8]:
                        if isinstance(val, str) and val.startswith("="):
                            ws.write_formula(r, c, val, colored_num_fmt)
                        else:
                            ws.write(r, c, val, colored_num_fmt)
                    elif isinstance(val, (int, float)):
                        ws.write(r, c, val, num_fmt)
                    elif isinstance(val, str) and val.startswith("="):
                        ws.write_formula(r, c, val, num_fmt)
                    else:
                        ws.write(r, c, val, text_fmt)
                elif xml_type == "Globe":
                    if c in [3, 5, 10]:
                        ws.write(r, c, val, colored_num_fmt)
                    elif 3 <= c <= 14 and c not in [3, 5, 10]:
                        ws.write(r, c, val, num_fmt)
                    elif c == 0:
                        ws.write(r, c, val, general_fmt)
                    elif c == 1:
                        ws.write(r, c, val, text_fmt)
                    else:
                        ws.write(r, c, val, num_fmt)
                else:
                    if c in [3, 5, 8]:
                        ws.write(r, c, val, colored_num_fmt)
                    elif 3 <= c <= 12 and c not in [3, 5, 8]:
                        ws.write(r, c, val, num_fmt)
                    elif c == 0:
                        ws.write(r, c, val, general_fmt)
                    elif c == 1:
                        ws.write(r, c, val, text_fmt)
                    else:
                        ws.write(r, c, val, num_fmt)

        for r in range(2, len(df)):
            cell_value = df.iloc[r, 0]
            if isinstance(cell_value, str):
                try:
                    hh, mm, ss = map(int, cell_value.strip().split(' ')[1].split(':'))
                    if ss != 0:
                        for c in range(df.shape[1]):
                            val = df.iloc[r, c]
                            if xml_type == "Den":
                                if c == 1:
                                    fmt_props = {'num_format': '@', 'border': 1, 'align': 'right'}
                                elif isinstance(val, str) and val.startswith("=") or isinstance(val, (int, float)):
                                    fmt_props = {'num_format': '0.00', 'border': 1, 'align': 'right'}
                                else:
                                    fmt_props = {'num_format': '@', 'border': 1, 'align': 'right'}
                            elif xml_type == "Globe":
                                if c in [3, 5, 10]:
                                    fmt_props = {'num_format': '0.00', 'border': 1, 'align': 'right', 'bg_color': '#B4C6E7'}
                                elif 3 <= c <= 14 and c not in [3, 5, 10]:
                                    fmt_props = {'num_format': '0.00', 'border': 1, 'align': 'right'}
                                elif c == 0:
                                    fmt_props = {'num_format': 'General', 'border': 1, 'align': 'right'}
                                elif c == 1:
                                    fmt_props = {'num_format': '@', 'border': 1, 'align': 'right'}
                                else:
                                    fmt_props = {'num_format': '0.00', 'border': 1, 'align': 'right'}
                            else:
                                if c in [3, 5, 8]:
                                    fmt_props = {'num_format': '0.00', 'border': 1, 'align': 'right', 'bg_color': '#FFFF00'}
                                elif 3 <= c <= 12 and c not in [3, 5, 8]:
                                    fmt_props = {'num_format': '0.00', 'border': 1, 'align': 'right'}
                                elif c == 0:
                                    fmt_props = {'num_format': 'General', 'border': 1, 'align': 'right'}
                                elif c == 1:
                                    fmt_props = {'num_format': '@', 'border': 1, 'align': 'right'}
                                else:
                                    fmt_props = {'num_format': '0.00', 'border': 1, 'align': 'right'}
                            fmt_props['bg_color'] = '#FFFF00'
                            highlight_fmt = workbook.add_format(fmt_props)
                            if isinstance(val, str) and val.startswith("="):
                                ws.write_formula(r, c, val, highlight_fmt)
                            else:
                                ws.write(r, c, val, highlight_fmt)
                except Exception:
                    continue

        if xml_type == "Den":
            widths = [17.73, 17.27] + [16.27] * 7 + [32.27, 36.36, 23.36, 24.76]
            hidden_cols = {6: 16.27}
        elif xml_type == "Globe":
            widths = [17.73, 17.27] + [14.91] * 9 + [41.91, 33.27, 43.36, 23.36]
            hidden_cols = {6: 14.91, 7: 14.91, 8: 14.91}
        else:
            widths = [17.73, 17.27] + [16.91] * 7 + [18.73, 17.55, 23.36, 23.36]
            hidden_cols = {6: 16.91}
        for i, w in enumerate(widths[:df.shape[1]]):
            ws.set_column(i, i, w)
        for col, w in hidden_cols.items():
            if col < df.shape[1]:
                ws.set_column(col, col, w, None, {'hidden': True})


class Worker(QObject):
    progress = pyqtSignal(int)
    error = pyqtSignal(str)
    dataReady = pyqtSignal(object, str, str)

    def __init__(self, xml_files, xml_type, format_definition=None):
        super().__init__()
        self.xml_files = xml_files
        self.xml_type = xml_type
        self.format_definition = format_definition
        self.current_file_index = 0
        self.cancel_requested = False

    @pyqtSlot()
    def request_cancel(self):
        self.cancel_requested = True

    @pyqtSlot()
    def process(self):
        while self.current_file_index < len(self.xml_files):
            if self.cancel_requested:
                self.error.emit("Operation cancelled by user.")
                return
            xml_file = self.xml_files[self.current_file_index]
            try:
                df_xml = pd.read_xml(
                    xml_file,
                    xpath=".//ns:Items",
                    namespaces={"ns": "http://tempuri.org/ArrayFieldDataSet.xsd"}
                )
            except Exception as e:
                self.error.emit(f"Error reading XML {xml_file}: {e}")
                self.current_file_index += 1
                continue
            self.progress.emit(50)
            try:
                if self.xml_type == "Den":
                    self.process_den_from_df(df_xml, xml_file)
                elif self.xml_type == "Globe":
                    self.process_globe_from_df(df_xml, xml_file)
                elif self.xml_type == "Glacier":
                    self.process_glacier_from_df(df_xml, xml_file)
                elif self.format_definition:
                    self.process_custom_from_df(df_xml, xml_file)
                else:
                    self.error.emit(f"Format '{self.xml_type}' is not implemented.")
                    self.current_file_index += 1
                    continue
                return
            except Exception as e:
                self.error.emit(f"An unexpected error occurred for {xml_file}: {e}")
                self.current_file_index += 1

    def _custom_label_presets(self):
        return [
            {"key": "clock", "row1": "0-0:1.0.0", "row2": "Clock"},
            {"key": "edis_status", "row1": "0-0:96.240.12 [hex]", "row2": "EDIS status"},
            {"key": "last_avg_demand", "row1": "1-1:1.5.0 [kW]", "row2": "Last average demand +A (QI+QIV)"},
            {"key": "demand", "row1": "", "row2": "Demand"},
            {"key": "active_import", "row1": "1-1:1.8.0 [Wh]", "row2": "Active energy import +A (QI+QIV)"},
            {"key": "kwh", "row1": "", "row2": "kWh"},
            {"key": "active_export", "row1": "1-1:2.8.0 [Wh]", "row2": "Active energy export -A (QII+QIII)"},
            {"key": "reactive_import", "row1": "1-1:3.8.0 [varh]", "row2": "Reactive energy import +R (QI+QII)"},
            {"key": "kvarh", "row1": "", "row2": "kVarh"},
            {"key": "reactive_export", "row1": "1-1:4.8.0 [varh]", "row2": "Reactive energy export -R (QIII+QIV)"},
        ]

    def _label_presets_map(self):
        return {item.get("key", ""): item for item in self._custom_label_presets() if isinstance(item, dict)}


    def process_custom_from_df(self, df, xml_file):
        try:
            df_filtered = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)
            if df_filtered.empty:
                self.error.emit(f"No valid rows found in {self.xml_type} XML.")
                return
            df_data = df_filtered.iloc[:, 1:]
            columns = self.format_definition.get("columns", [])
            if not columns:
                self.error.emit(f"Format '{self.xml_type}' has no columns.")
                return

            max_col = max(1, len(columns))
            formula_columns = [
                idx for idx, col in enumerate(columns)
                if str(col.get("type", "empty")) == "formula"
            ]
            label_map = self._label_presets_map()
            header_row_1 = [""] * max_col
            header_row_2 = [""] * max_col
            final_rows = []
            for i in range(len(df_data)):
                row = [""] * max_col
                                                                          
                excel_row_num = i + 3
                for target, col_def in enumerate(columns):
                    if 0 <= target < max_col:
                        label_key = str(col_def.get("labelKey", "")).strip()
                        if label_key.startswith("custom:"):
                            header_row_1[target] = ""
                            header_row_2[target] = label_key[7:].strip()
                        else:
                            preset = label_map.get(label_key)
                            if preset:
                                header_row_1[target] = str(preset.get("row1", ""))
                                header_row_2[target] = str(preset.get("row2", ""))
                    col_type = str(col_def.get("type", "empty"))
                    value = str(col_def.get("value", ""))
                    if col_type == "data":
                        try:
                            source_idx = int(value)
                        except Exception:
                            source_idx = -1
                        if 0 <= source_idx < df_data.shape[1] and 0 <= target < max_col:
                            row[target] = df_data.iloc[i, source_idx]
                    elif col_type == "formula":
                        if 0 <= target < max_col:
                            row[target] = value.replace("{r}", str(excel_row_num)).replace("{r-1}", str(max(1, excel_row_num - 1)))
                final_rows.append(row)

            final_df = pd.DataFrame(final_rows)
            final_df.attrs["custom_widths"] = [float(col.get("width", 14) or 14) for col in columns]
            final_df.attrs["formula_columns"] = [i for i in formula_columns if 0 <= i < max_col]
            final_df.attrs["custom_header_row_1"] = header_row_1
            final_df.attrs["custom_header_row_2"] = header_row_2
            final_df.attrs["data_start_row"] = 3
            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type, xml_file)
        except Exception as e:
            self.error.emit(f"{self.xml_type} processing error: {e}")

    def process_den_from_df(self, df, xml_file):
        try:
            df_filtered = df[~df.iloc[:,0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)
            if df_filtered.empty:
                self.error.emit("No valid rows found in Den XML.")
                return
            df_data = df_filtered.iloc[:,1:12]
            column_mapping = {0:0,1:1,2:2,3:4,4:6,5:7,6:9,7:10,8:11,9:12}
            max_col = max(column_mapping.values())+1
            header_den_row_1 = [""]*max_col
            header_den_row_2 = [""]*max_col
            for idx,val in {0:"0-0:1.0.0",1:"0-0:96.240.12 [hex]",2:"1-1:1.5.0 [kW]",4:"1-1:1.8.0 [Wh]",6:"1-1:2.8.0 [Wh]",7:"1-1:3.8.0 [varh]",9:"1-1:4.8.0 [varh]",10:"1-1:15.8.1 [Wh]",11:"1-1:13.5.0",12:"1-1:128.8.0 [Wh]"}.items(): header_den_row_1[idx]=val
            for idx,val in {0:"Clock",1:"EDIS status",2:"Last average demand +A (QI+QIV)",3:"Demand",4:"Active energy import +A (QI+QIV)",5:"kWh",6:"Active energy export -A (QII+QIII)",7:"Reactive energy import +R (QI+QII)",8:"kVarh",9:"Reactive energy export -R (QIII+QIV)",10:"Active energy A (QI+QII+QIII+QIV) rate 1",11:"Last average power factor",12:"Energy |AL1|+|AL2|+|AL3|"}.items(): header_den_row_2[idx]=val
            final_rows = [header_den_row_1, header_den_row_2]
            for i in range(len(df_data)):
                row = [""]*max_col
                for idx,col in column_mapping.items(): row[col]=df_data.iloc[i,idx]
                if i>0:
                    row[3]=f"=C{i+3}*280"; row[5]=f"=(E{i+3}-E{i+2})*280/1000"; row[8]=f"=(H{i+3}-H{i+2})*280/1000"
                final_rows.append(row)
            final_df = pd.DataFrame(final_rows)
            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type, xml_file)
        except Exception as e:
            self.error.emit(f"Den processing error: {e}")

    def process_globe_from_df(self, df, xml_file):
        try:
            df_filtered = df[~df.iloc[:,0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)
            if df_filtered.empty:
                self.error.emit("No valid rows found in Globe XML.")
                return
            df_data = df_filtered.iloc[:,1:14]
            column_mapping = {0:0,1:1,2:2,3:4,4:6,5:7,6:8,7:9,8:11,9:12,10:13,11:14}
            max_col = max(column_mapping.values())+1
            header_globe_row_1 = [""]*max_col
            header_globe_row_2 = [""]*max_col
            for idx,val in {0:"0-0:1.0.0",1:"0-0:96.240.12 [hex]",2:"1-1:1.5.0 [kW]",4:"1-1:1.8.0 [Wh]",6:"1-1:1.29.0 [Wh]",7:"1-1:2.8.0 [Wh]",8:"1-1:2.29.0 [Wh]",9:"1-1:3.8.0 [varh]",11:"1-1:3.29.0 [varh]",12:"1-1:4.8.0 [varh]",13:"1-1:4.29.0 [varh]",14:"1-1:13.5.0"}.items(): header_globe_row_1[idx]=val
            for idx,val in {0:"Clock",1:"EDIS status",2:"Last average demand +A (QI+QIV)",3:"Demand",4:"Active energy import +A (QI+QIV)",5:"kWh",6:"Energy delta over capture period 1 +A (QI+QIV)",7:"Active energy export -A (QII+QIII)",8:"Energy delta over capture period 1 -A (QII+QIII)",9:"Reactive energy import +R (QI+QII)",10:"kVarh",11:"Energy delta over capture period 1 +R (QI+QII)",12:"Reactive energy export -R (QIII+QIV)",13:"Energy delta over capture period 1 -R (QIII+QIV)",14:"Last average power factor"}.items(): header_globe_row_2[idx]=val
            final_rows = [header_globe_row_1, header_globe_row_2]
            total_rows = len(df_data)
            for i in range(total_rows):
                row = [""]*max_col
                for idx,col in column_mapping.items(): row[col]=df_data.iloc[i,idx]
                if i>0: row[3]=f"=C{i+3}*1400"; row[5]=f"=(E{i+3}-E{i+2})*1400/1000"; row[10]=f"=(J{i+3}-J{i+2})*1400/1000"
                final_rows.append(row)
                self.progress.emit(50+int((i/total_rows)*40))
            final_df = pd.DataFrame(final_rows)
            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type, xml_file)
        except Exception as e:
            self.error.emit(f"Globe processing error: {e}")

    def process_glacier_from_df(self, df, xml_file):
        try:
            df_filtered = df[~df.iloc[:,0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)
            if df_filtered.empty:
                self.error.emit("No valid rows found in Glacier XML.")
                return
            df_data = df_filtered.iloc[:,1:12]
            column_mapping = {0:0,1:1,2:2,3:4,4:6,5:7,6:9,7:10,8:11,9:12}
            max_col = max(column_mapping.values())+1
            header_row_1 = [""]*max_col
            header_row_2 = [""]*max_col
            for idx,val in {0:"0-0:1.0.0",1:"0-0:96.240.12 [hex]",2:"1-1:1.5.0 [kW]",4:"1-1:1.8.0 [Wh]",6:"1-1:2.8.0 [Wh]",7:"1-1:3.8.0 [varh]",9:"1-1:4.8.0 [varh]",10:"1-1:15.8.1 [Wh]",11:"1-1:13.5.0",12:"1-1:128.8.0 [Wh]"}.items(): header_row_1[idx]=val
            for idx,val in {0:"Clock",1:"EDIS status",2:"Last average demand +A (QI+QIV)",3:"Demand",4:"Active energy import +A (QI+QIV)",5:"kWh",6:"Active energy export -A (QII+QIII)",7:"Reactive energy import +R (QI+QII)",8:"kVarh",9:"Reactive energy export -R (QIII+QIV)",10:"Active energy A (QI+QII+QIII+QIV) rate 1",11:"Last average power factor",12:"Energy |AL1|+|AL2|+|AL3|"}.items(): header_row_2[idx]=val
            final_rows = [header_row_1, header_row_2]
            for i in range(len(df_data)):
                row = [""]*max_col
                for idx,col in column_mapping.items(): row[col]=df_data.iloc[i,idx]
                if i>0: row[3]=f"=C{i+3}*280"; row[5]=f"=(E{i+3}-E{i+2})*280/1000"; row[8]=f"=(H{i+3}-H{i+2})*280/1000"
                final_rows.append(row)
            final_df = pd.DataFrame(final_rows)
            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type, xml_file)
        except Exception as e:
            self.error.emit(f"Glacier processing error: {e}")


class PathDiscoveryWorker(QObject):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, paths):
        super().__init__()
        self.paths = [normalize_path(p) for p in paths if normalize_path(p)]
        self.cancel_requested = False

    @pyqtSlot()
    def request_cancel(self):
        self.cancel_requested = True

    @pyqtSlot()
    def process(self):
        try:
            xml_paths = collect_xml_files_from_paths(self.paths, should_stop=lambda: self.cancel_requested)
            if self.cancel_requested:
                self.error.emit("Operation cancelled by user.")
                return
            self.finished.emit(xml_paths)
        except Exception as e:
            self.error.emit(f"Failed to scan dropped paths: {e}")

class SaveWorker(QObject):
    progress = pyqtSignal(int)
    error = pyqtSignal(str)
    saved = pyqtSignal(str, str)

    def __init__(self, df, xml_type, save_path, xml_file):
        super().__init__()
        self.df = df
        self.xml_type = xml_type
        self.save_path = save_path
        self.xml_file = xml_file
        self.cancel_requested = False

    @pyqtSlot()
    def request_cancel(self):
        self.cancel_requested = True

    @pyqtSlot()
    def save(self):
        try:
            if self.cancel_requested:
                self.error.emit("Operation cancelled by user.")
                return
            self.progress.emit(90)
            export_dataframe_to_excel(self.df, self.xml_type, self.save_path, self.xml_file)
            self.progress.emit(100)
            self.saved.emit(self.save_path, self.xml_file)
        except PermissionError as e:
            self.error.emit(f"PERMISSION_DENIED::{self.save_path}::{e}")
        except Exception as e:
            self.error.emit(f"Failed to save Excel: {e}")
            


class BatchSaveWorker(QObject):
    progress = pyqtSignal(int)
    error = pyqtSignal(str)
    finished = pyqtSignal(object)

    def __init__(self, batch_results, batch_outputs):
        super().__init__()
        self.batch_results = batch_results
        self.batch_outputs = [dict(item) for item in batch_outputs]
        self.cancel_requested = False

    @pyqtSlot()
    def request_cancel(self):
        self.cancel_requested = True

    @pyqtSlot()
    def save_all(self):
        try:
            total = len(self.batch_results)
            if total == 0:
                self.finished.emit(self.batch_outputs)
                return

            for i, result in enumerate(self.batch_results):
                if self.cancel_requested:
                    self.error.emit("Operation cancelled by user.")
                    return
                output = self.batch_outputs[i]
                save_path = os.path.join(output["saveDir"], ensure_xlsx_extension(output["fileName"]))
                try:
                    export_dataframe_to_excel(result["df"], result["xml_type"], save_path, result["xml_file"])
                except PermissionError as e:
                    self.error.emit(f"BATCH_PERMISSION_DENIED::{i}::{save_path}::{e}")
                    return
                self.batch_outputs[i]["savePath"] = save_path
                self.batch_outputs[i]["fileName"] = os.path.basename(save_path)
                self.progress.emit(int(((i + 1) / total) * 100))

            self.finished.emit(self.batch_outputs)
        except Exception as e:
            self.error.emit(f"Failed to save batch files: {e}")


class Backend(QObject):
    progressUpdated = pyqtSignal(int)
    formatModelChanged = pyqtSignal()
    formatDesignerStatusChanged = pyqtSignal()
    xmlTypeOptionsChanged = pyqtSignal()
    formatSavePathChanged = pyqtSignal()
    xmlPreviewChanged = pyqtSignal()

    def __init__(self, engine):
        super().__init__()
        self.engine = engine
        self.root = None
        self.selected_file = None
        self.selected_files = []
        self.xml_type = ""
        self.progress = 0
        self.is_batch = False
        self.current_batch_index = 0
        self.progressUpdated.connect(self.updateProgressInQML)
        self.thread = None
        self.worker = None
        self.save_thread = None
        self.save_worker = None
        self.batch_save_thread = None
        self.batch_save_worker = None
        self.path_scan_thread = None
        self.path_scan_worker = None
        self.batch_results = []
        self.batch_outputs = []
        self.batch_file_statuses = []
        self.cancel_requested = False
        self.formats_dir = _initial_formats_dir()
        self.formats_path = os.path.join(self.formats_dir, "format_model.json")
        self.format_save_path = self.formats_path
        self.custom_label_options = self._custom_label_presets()
        self.format_model = self._load_or_default_formats()
        self.xml_type_options = []
        self.format_designer_status = ""
        self.xml_preview_headers = []
        self.xml_preview_rows = []
        self.xml_preview_status = ""
        self.preview_selected_file = ""
        self._format_edit_snapshot = None
        self._format_edit_active = False
        self._last_save_payload = None
        self.settings = QSettings("CubeFlow", "CubeFlow")
        self.last_open_dir = str(self.settings.value("lastOpenDir", "", str))
        self.last_save_dir = str(self.settings.value("lastSaveDir", "", str))
        self.last_batch_dir = str(self.settings.value("lastBatchDir", "", str))
        self.xml_type = str(self.settings.value("lastXmlType", "", str))
        saved_format_path = str(self.settings.value("formatSavePath", "", str)).strip()
        if saved_format_path and "AppData\\Local\\CubeFlow\\formats\\format_model.json" in saved_format_path:
            self.settings.remove("formatSavePath")
        self.format_save_path = self.formats_path
        self._refresh_xml_type_options(emit_signal=False)

    @pyqtProperty('QVariantList', notify=formatModelChanged)
    def formatModel(self):
        return self.format_model

    @pyqtProperty(str, notify=formatDesignerStatusChanged)
    def formatDesignerStatus(self):
        return self.format_designer_status

    @pyqtProperty('QVariantList', notify=xmlTypeOptionsChanged)
    def xmlTypeOptions(self):
        return self.xml_type_options

    @pyqtProperty('QVariantList', notify=formatModelChanged)
    def customLabelOptions(self):
        return self.custom_label_options

    @pyqtProperty(str, notify=formatSavePathChanged)
    def formatSavePath(self):
        return self.format_save_path

    @pyqtProperty('QVariantList', notify=xmlPreviewChanged)
    def xmlPreviewHeaders(self):
        return self.xml_preview_headers

    @pyqtProperty('QVariantList', notify=xmlPreviewChanged)
    def xmlPreviewRows(self):
        return self.xml_preview_rows

    @pyqtProperty(str, notify=xmlPreviewChanged)
    def xmlPreviewStatus(self):
        return self.xml_preview_status

    def _set_format_designer_status(self, status):
        self.format_designer_status = status
        self.formatDesignerStatusChanged.emit()

    def _set_xml_preview(self, headers=None, rows=None, status=""):
        self.xml_preview_headers = headers if isinstance(headers, list) else []
        self.xml_preview_rows = rows if isinstance(rows, list) else []
        self.xml_preview_status = str(status or "")
        self.xmlPreviewChanged.emit()

    def _preview_source_file(self):
        if self.preview_selected_file and os.path.exists(self.preview_selected_file):
            return self.preview_selected_file
        if self.selected_file and os.path.exists(self.selected_file):
            return self.selected_file
        if self.selected_files:
            first = self.selected_files[0]
            if first and os.path.exists(first):
                return first
        return ""

    def _default_columns(self, formula):
        return [
            {"col": "A", "type": "data", "value": "0", "width": 17, "labelKey": ""},
            {"col": "B", "type": "data", "value": "1", "width": 17, "labelKey": ""},
            {"col": "C", "type": "data", "value": "2", "width": 14, "labelKey": ""},
            {"col": "D", "type": "formula", "value": formula, "width": 14, "labelKey": ""},
        ]

    def _build_columns_from_spec(self, max_col, mapping, formulas, widths, label_keys=None):
        target_to_source = {target: source for source, target in mapping.items()}
        label_keys = label_keys if isinstance(label_keys, dict) else {}
        columns = []
        for idx in range(max_col):
            if idx in formulas:
                col_type = "formula"
                value = formulas[idx]
            elif idx in target_to_source:
                col_type = "data"
                value = str(target_to_source[idx])
            else:
                col_type = "empty"
                value = ""
            width = widths[idx] if idx < len(widths) else 14
            columns.append({
                "col": index_to_excel_col(idx),
                "type": col_type,
                "value": value,
                "width": width,
                "labelKey": str(label_keys.get(idx, "")),
            })
        return columns

    def _custom_label_presets(self):
        return [
            {"key": "clock", "row1": "0-0:1.0.0", "row2": "Clock"},
            {"key": "edis_status", "row1": "0-0:96.240.12 [hex]", "row2": "EDIS status"},
            {"key": "last_avg_demand", "row1": "1-1:1.5.0 [kW]", "row2": "Last average demand +A (QI+QIV)"},
            {"key": "demand", "row1": "", "row2": "Demand"},
            {"key": "active_import", "row1": "1-1:1.8.0 [Wh]", "row2": "Active energy import +A (QI+QIV)"},
            {"key": "kwh", "row1": "", "row2": "kWh"},
            {"key": "active_export", "row1": "1-1:2.8.0 [Wh]", "row2": "Active energy export -A (QII+QIII)"},
            {"key": "reactive_import", "row1": "1-1:3.8.0 [varh]", "row2": "Reactive energy import +R (QI+QII)"},
            {"key": "kvarh", "row1": "", "row2": "kVarh"},
            {"key": "reactive_export", "row1": "1-1:4.8.0 [varh]", "row2": "Reactive energy export -R (QIII+QIV)"},
        ]

    def _label_presets_map(self):
        return {item.get("key", ""): item for item in self.custom_label_options if isinstance(item, dict)}

    def _default_formats(self):
        den_mapping = {0: 0, 1: 1, 2: 2, 3: 4, 4: 6, 5: 7, 6: 9, 7: 10, 8: 11, 9: 12}
        den_formulas = {
            3: "=C{r}*280",
            5: "=(E{r}-E{r-1})*280/1000",
            8: "=(H{r}-H{r-1})*280/1000",
        }
        den_label_keys = {
            0: "clock",
            1: "edis_status",
            2: "last_avg_demand",
            3: "demand",
            4: "active_import",
            5: "kwh",
            6: "active_export",
            7: "reactive_import",
            8: "kvarh",
            9: "reactive_export",
        }
        den_widths = [17.73, 17.27] + [16.27] * 7 + [32.27, 36.36, 23.36, 24.76]

        globe_mapping = {0: 0, 1: 1, 2: 2, 3: 4, 4: 6, 5: 7, 6: 8, 7: 9, 8: 11, 9: 12, 10: 13, 11: 14}
        globe_formulas = {
            3: "=C{r}*1400",
            5: "=(E{r}-E{r-1})*1400/1000",
            10: "=(J{r}-J{r-1})*1400/1000",
        }
        globe_label_keys = {
            0: "clock",
            1: "edis_status",
            2: "last_avg_demand",
            3: "demand",
            4: "active_import",
            5: "kwh",
            7: "active_export",
            9: "reactive_import",
            10: "kvarh",
            12: "reactive_export",
        }
        globe_widths = [17.73, 17.27] + [14.91] * 9 + [41.91, 33.27, 43.36, 23.36]

        glacier_mapping = den_mapping
        glacier_formulas = den_formulas
        glacier_label_keys = den_label_keys
        glacier_widths = [17.73, 17.27] + [16.91] * 7 + [18.73, 17.55, 23.36, 23.36]

        return [
            {"name": "Den", "columns": self._build_columns_from_spec(13, den_mapping, den_formulas, den_widths, den_label_keys)},
            {"name": "Glacier", "columns": self._build_columns_from_spec(13, glacier_mapping, glacier_formulas, glacier_widths, glacier_label_keys)},
            {"name": "Globe", "columns": self._build_columns_from_spec(15, globe_mapping, globe_formulas, globe_widths, globe_label_keys)},
        ]

    def _normalize_loaded_formats(self, raw_formats):
        normalized = []
        if isinstance(raw_formats, dict):
            raw_formats = [raw_formats]
        if not isinstance(raw_formats, list):
            return normalized
        for item in raw_formats:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            if not name:
                continue
            raw_columns = item.get("columns", [])
            columns = []
            if isinstance(raw_columns, list):
                for col in raw_columns:
                    if not isinstance(col, dict):
                        continue
                    row_type = self._sanitize_format_type(col.get("type", "data"))
                    columns.append({
                        "col": self._normalize_column_label(col.get("col", "A")),
                        "type": row_type,
                        "value": self._sanitize_format_value(row_type, col.get("value", "")),
                        "width": self._sanitize_format_width(col.get("width", 14)),
                        "labelKey": self._sanitize_label_key(col.get("labelKey", ""), row_type),
                    })
            if not columns:
                columns = self._default_columns("=C{r}*280")
            normalized.append({"name": name, "columns": columns})
        return normalized

    def _merge_format_entries(self, base_formats, extra_formats):
        merged = []
        seen = set()
        for fmt in base_formats:
            name = str((fmt or {}).get("name", "")).strip()
            if not name:
                continue
            key = name.lower()
            if key in seen:
                continue
            merged.append(fmt)
            seen.add(key)
        for fmt in extra_formats:
            name = str((fmt or {}).get("name", "")).strip()
            if not name:
                continue
            key = name.lower()
            if key in seen:
                continue
            merged.append(fmt)
            seen.add(key)
        return merged

    def _load_sidecar_formats(self):
        extras = []
        try:
            if not os.path.isdir(self.formats_dir):
                return extras
            for entry in os.listdir(self.formats_dir):
                if not entry.lower().endswith(".json"):
                    continue
                if entry.lower() == "format_model.json":
                    continue
                path = os.path.join(self.formats_dir, entry)
                if not os.path.isfile(path):
                    continue
                try:
                    with open(path, "r", encoding="utf-8") as fp:
                        loaded = json.load(fp)
                    parsed = self._normalize_loaded_formats(loaded)
                    if parsed:
                        extras.extend(parsed)
                except Exception:
                    continue
        except Exception:
            return extras
        return extras

    def _load_or_default_formats(self):
        self._ensure_formats_storage_writable()
        os.makedirs(self.formats_dir, exist_ok=True)
        base_formats = []
        if os.path.exists(self.formats_path):
            try:
                with open(self.formats_path, "r", encoding="utf-8") as fp:
                    loaded = json.load(fp)
                parsed = self._normalize_loaded_formats(loaded)
                if parsed:
                    base_formats = parsed
            except Exception:
                pass
        if not base_formats:
            bundled_formats_path = os.path.join(RESOURCE_BASE_DIR, "formats", "format_model.json")
            if os.path.exists(bundled_formats_path):
                try:
                    with open(bundled_formats_path, "r", encoding="utf-8") as fp:
                        loaded = json.load(fp)
                    parsed = self._normalize_loaded_formats(loaded)
                    if parsed:
                        base_formats = parsed
                except Exception:
                    pass
        if not base_formats:
            base_formats = self._default_formats()
        sidecars = self._load_sidecar_formats()
        return self._merge_format_entries(base_formats, sidecars)

    def _refresh_xml_type_options(self, emit_signal=True):
        self.xml_type_options = [fmt.get("name", "") for fmt in self.format_model if fmt.get("name", "")]
        if emit_signal:
            self.xmlTypeOptionsChanged.emit()

    def _unique_format_name(self, base_name, skip_index=None):
        raw = (base_name or "").strip()
        if not raw:
            raw = "New Format"
        existing = {
            self.format_model[i]["name"].lower()
            for i in range(len(self.format_model))
            if i != skip_index
        }
        if raw.lower() not in existing:
            return raw
        counter = 2
        while True:
            candidate = f"{raw} {counter}"
            if candidate.lower() not in existing:
                return candidate
            counter += 1

    def _next_column_label(self, columns):
        used_indexes = set()
        if isinstance(columns, list):
            for row in columns:
                col_label = self._normalize_column_label((row or {}).get("col", "A"))
                col_index = excel_col_to_index(col_label)
                if col_index >= 0:
                    used_indexes.add(col_index)
        next_index = 0
        while next_index in used_indexes:
            next_index += 1
        return index_to_excel_col(next_index)

    def _normalize_column_label(self, value):
        normalized = "".join(ch for ch in str(value).upper() if ch.isalpha())
        return normalized[:3] if normalized else "A"

    def _sanitize_column_input(self, value):
        return "".join(ch for ch in str(value).upper() if ch.isalpha())[:3]

    def _sanitize_format_type(self, value):
        type_value = str(value).strip().lower()
        return type_value if type_value in ("data", "formula", "empty") else "data"

    def _sanitize_format_value(self, row_type, value):
        safe_type = self._sanitize_format_type(row_type)
        if safe_type == "empty":
            return ""
        return str(value)

    def _sanitize_format_width(self, value):
        try:
            width_value = int(value)
        except Exception:
            width_value = 14
        return max(1, min(200, width_value))

    def _allowed_label_keys_for_type(self, row_type):
        safe_type = self._sanitize_format_type(row_type)
        all_keys = {item.get("key", "") for item in self.custom_label_options if isinstance(item, dict)}
        if safe_type == "formula":
            return {"demand", "kwh", "kvarh"}
        return all_keys

    def _sanitize_label_key(self, value, row_type="data"):
        key = str(value or "").strip()
        if key.startswith("custom:"):
            custom_text = key[7:].strip()
            return f"custom:{custom_text}" if custom_text else ""
        valid_keys = self._allowed_label_keys_for_type(row_type)
        return key if key in valid_keys else ""

    def _sort_format_columns(self, columns):
        if not isinstance(columns, list) or not columns:
            return
        columns.sort(key=lambda row: excel_col_to_index(self._normalize_column_label((row or {}).get("col", "A"))))

    def _is_builtin_format_name(self, name):
        return str(name).strip().lower() in ("den", "glacier", "globe")

    def _only_builtin_formats_left(self):
        if not self.format_model:
            return True
        for fmt in self.format_model:
            if not self._is_builtin_format_name(fmt.get("name", "")):
                return False
        return True

    def _persist_formats_after_delete(self):
        target_path = self.formats_path
        try:
            os.makedirs(self.formats_dir, exist_ok=True)
            if self._only_builtin_formats_left():
                if os.path.exists(target_path):
                    self._prepare_file_for_write(target_path)
                    with open(target_path, "w", encoding="utf-8") as fp:
                        json.dump(self.format_model, fp, indent=2)
                    self._hide_file_if_supported(target_path)
                    self._set_format_designer_status(f"Kept format file and saved built-in formats: {target_path}")
                else:
                    self._set_format_designer_status("Only built-in formats remain. No format file found to delete.")
                return
            self._prepare_file_for_write(target_path)
            with open(target_path, "w", encoding="utf-8") as fp:
                json.dump(self.format_model, fp, indent=2)
            self._hide_file_if_supported(target_path)
            self._set_format_designer_status(f"Updated format file: {target_path}")
        except Exception as e:
            self._set_format_designer_status(f"Failed to update format file: {e}")

    def _autosave_formats(self):
        if not self._ensure_formats_storage_writable():
            self._set_format_designer_status("Failed to auto-save format file: no writable format storage path.")
            return
        try:
            os.makedirs(self.formats_dir, exist_ok=True)
            self._prepare_file_for_write(self.formats_path)
            with open(self.formats_path, "w", encoding="utf-8") as fp:
                json.dump(self.format_model, fp, indent=2)
            self._hide_file_if_supported(self.formats_path)
            return
        except Exception:
            pass
        if self._switch_to_next_writable_formats_storage():
            try:
                os.makedirs(self.formats_dir, exist_ok=True)
                self._prepare_file_for_write(self.formats_path)
                with open(self.formats_path, "w", encoding="utf-8") as fp:
                    json.dump(self.format_model, fp, indent=2)
                self._hide_file_if_supported(self.formats_path)
                return
            except Exception as e:
                self._set_format_designer_status(f"Failed to auto-save format file: {e}")
                return
        self._set_format_designer_status("Failed to auto-save format file: permission denied for all storage paths.")

    def _probe_writable_dir(self, directory):
        try:
            os.makedirs(directory, exist_ok=True)
            probe = os.path.join(directory, ".write_probe.tmp")
            with open(probe, "w", encoding="utf-8") as fp:
                fp.write("ok")
            os.remove(probe)
            return True
        except Exception:
            return False

    def _format_storage_candidates(self):
        current = self.formats_dir
        candidates = [current]
        docs = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        if docs:
            candidates.append(os.path.join(docs, "CubeFlow", "formats"))
        candidates.append(os.path.join(os.path.expanduser("~"), "Documents", "CubeFlow", "formats"))
        candidates.append(os.path.join(tempfile.gettempdir(), "CubeFlow", "formats"))
        unique = []
        seen = set()
        for item in candidates:
            key = os.path.normcase(os.path.normpath(item))
            if key in seen:
                continue
            seen.add(key)
            unique.append(item)
        return unique

    def _set_formats_storage(self, directory):
        self.formats_dir = directory
        self.formats_path = os.path.join(self.formats_dir, "format_model.json")
        self.format_save_path = self.formats_path
        if hasattr(self, "settings") and self.settings is not None:
            self.settings.setValue("formatSavePath", self.formats_path)
            self.formatSavePathChanged.emit()

    def _ensure_formats_storage_writable(self):
        for candidate in self._format_storage_candidates():
            if self._probe_writable_dir(candidate):
                if os.path.normcase(os.path.normpath(candidate)) != os.path.normcase(os.path.normpath(self.formats_dir)):
                    self._set_formats_storage(candidate)
                    self._set_format_designer_status(f"Switched format storage to writable path: {self.formats_path}")
                return True
        return False

    def _switch_to_next_writable_formats_storage(self):
        current_key = os.path.normcase(os.path.normpath(self.formats_dir))
        candidates = self._format_storage_candidates()
        start_idx = 0
        for i, c in enumerate(candidates):
            if os.path.normcase(os.path.normpath(c)) == current_key:
                start_idx = i + 1
                break
        for candidate in candidates[start_idx:]:
            if self._probe_writable_dir(candidate):
                self._set_formats_storage(candidate)
                self._set_format_designer_status(f"Switched format storage to writable path: {self.formats_path}")
                return True
        return False

    def _hide_file_if_supported(self, path):
        if not path or not os.path.exists(path):
            return
        if os.name != "nt":
            return
        try:
            import ctypes
            FILE_ATTRIBUTE_HIDDEN = 0x02
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
            if attrs == -1:
                return
            if not (attrs & FILE_ATTRIBUTE_HIDDEN):
                ctypes.windll.kernel32.SetFileAttributesW(str(path), attrs | FILE_ATTRIBUTE_HIDDEN)
        except Exception:
            pass

    def _prepare_file_for_write(self, path):
        if not path or os.name != "nt" or not os.path.exists(path):
            return
        try:
            import ctypes
            FILE_ATTRIBUTE_READONLY = 0x01
            FILE_ATTRIBUTE_HIDDEN = 0x02
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
            if attrs == -1:
                return
            new_attrs = attrs & ~FILE_ATTRIBUTE_READONLY
            new_attrs = new_attrs & ~FILE_ATTRIBUTE_HIDDEN
            ctypes.windll.kernel32.SetFileAttributesW(str(path), new_attrs)
        except Exception:
            pass

    def _autosave_format_model_changes(self):
        self._autosave_formats()

    def _safe_format_filename(self, name):
        raw = str(name or "").strip()
        if not raw:
            raw = "format"
        cleaned = "".join(ch for ch in raw if ch not in '<>:"/\\|?*' and ord(ch) >= 32).strip().rstrip(". ")
        if not cleaned:
            cleaned = "format"
        return f"{cleaned}.json"

    def _read_format_name_from_file(self, path):
        try:
            with open(path, "r", encoding="utf-8") as fp:
                data = json.load(fp)
            if isinstance(data, dict):
                return str(data.get("name", "")).strip()
        except Exception:
            pass
        return ""

    def _resolve_unique_format_file_path(self, format_name):
        base_file = self._safe_format_filename(format_name)
        base_stem, _ = os.path.splitext(base_file)
        candidate = os.path.join(self.formats_dir, base_file)
        counter = 2
        while os.path.exists(candidate):
            existing_name = self._read_format_name_from_file(candidate)
            if existing_name and existing_name.lower() == str(format_name).strip().lower():
                return candidate
            candidate = os.path.join(self.formats_dir, f"{base_stem} {counter}.json")
            counter += 1
        return candidate

    def _find_format_file_paths_by_name(self, name):
        matches = []
        target_name = str(name or "").strip().lower()
        if not target_name:
            return matches
        canonical = self._safe_format_filename(name).lower()
        try:
            for entry in os.listdir(self.formats_dir):
                if not entry.lower().endswith(".json"):
                    continue
                if entry.lower() == "format_model.json":
                    continue
                path = os.path.join(self.formats_dir, entry)
                if not os.path.isfile(path):
                    continue
                if entry.lower() == canonical:
                    matches.append(path)
                    continue
                existing_name = self._read_format_name_from_file(path).lower()
                if existing_name == target_name:
                    matches.append(path)
        except Exception:
            pass
        return matches

    def _delete_format_files_for_names(self, names):
        removed = []
        unique_paths = set()
        for name in names:
            for path in self._find_format_file_paths_by_name(name):
                unique_paths.add(path)
        for path in sorted(unique_paths):
            os.remove(path)
            removed.append(path)
        return removed

    def _delete_format_file_by_name(self, name):
        file_name = self._safe_format_filename(name)
        target_path = os.path.join(self.formats_dir, file_name)
        if os.path.exists(target_path):
            os.remove(target_path)
            return target_path
        return ""

    def refreshBatchFileStatusesProperty(self):
        if self.root:
            self.root.setProperty("batchFileStatuses", self.batch_file_statuses)

    def _stop_thread(self, thread_attr, worker_attr=None, timeout_ms=3000):
        t = getattr(self, thread_attr, None)
        if t:
            try:
                if t.isRunning():
                    t.quit()
                    if not t.wait(timeout_ms):
                        t.terminate()
                        t.wait(timeout_ms)
            except Exception:
                pass
        setattr(self, thread_attr, None)
        if worker_attr:
            setattr(self, worker_attr, None)

    def _request_worker_cancel(self, worker_attr):
        worker = getattr(self, worker_attr, None)
        if worker and hasattr(worker, "request_cancel"):
            try:
                worker.request_cancel()
            except Exception:
                pass

    def _request_all_worker_cancels(self):
        self._request_worker_cancel("worker")
        self._request_worker_cancel("save_worker")
        self._request_worker_cancel("batch_save_worker")
        self._request_worker_cancel("path_scan_worker")

    def _stop_all_background_threads(self, timeout_ms=3000):
        self._stop_thread("thread", "worker", timeout_ms)
        self._stop_thread("save_thread", "save_worker", timeout_ms)
        self._stop_thread("batch_save_thread", "batch_save_worker", timeout_ms)
        self._stop_thread("path_scan_thread", "path_scan_worker", timeout_ms)

    def applyRememberedSettingsToUI(self):
        if not self.root:
            return
        if self.xml_type:
            self.root.setProperty("selectionType", self.xml_type)

    @pyqtSlot()
    def openFormatDesigner(self):
        if self.root:
            self.root.setProperty("processState", "formatDesigner")

    @pyqtSlot()
    def closeFormatDesigner(self):
        if self.root:
            self.root.setProperty("processState", "idle")

    @pyqtSlot()
    def chooseFormatSavePath(self):
        start_dir = self.formats_dir
        chosen_path, _ = QFileDialog.getSaveFileName(
            None,
            "Save Formats As",
            start_dir,
            "JSON Files (*.json)"
        )
        if not chosen_path:
            return
        if not chosen_path.lower().endswith(".json"):
            chosen_path += ".json"
        self.format_save_path = self.formats_path
        self.settings.setValue("formatSavePath", self.formats_path)
        self.formatSavePathChanged.emit()
        self._set_format_designer_status(f"Formats are saved to {self.formats_path}")

    @pyqtSlot()
    def importFormatModelFromFile(self):
        start_dir = self.formats_dir
        chosen_path, _ = QFileDialog.getOpenFileName(
            None,
            "Open Format File",
            start_dir,
            "JSON Files (*.json)"
        )
        if not chosen_path:
            return
        try:
            with open(chosen_path, "r", encoding="utf-8") as fp:
                loaded = json.load(fp)
            if isinstance(loaded, dict):
                loaded = [loaded]
            parsed = self._normalize_loaded_formats(loaded)
            if not parsed:
                QMessageBox.warning(None, "Import Failed", "No valid format entries were found in the selected JSON file.")
                self._set_format_designer_status("Failed to import: no valid format entries found.")
                return

            added = 0
            for fmt in parsed:
                name = self._unique_format_name(fmt.get("name", "New Format"))
                self.format_model.append({
                    "name": name,
                    "columns": fmt.get("columns", self._default_columns("=C{r}*280"))
                })
                added += 1

            self.format_save_path = self.formats_path
            self.settings.setValue("formatSavePath", self.formats_path)
            self.formatSavePathChanged.emit()
            self.formatModelChanged.emit()
            self._refresh_xml_type_options()
            self._autosave_formats()
            self._set_format_designer_status(f"Imported {added} format(s) from {chosen_path}")
            QMessageBox.information(None, "Formats Imported", f"Imported {added} format(s) from:\n{chosen_path}")
        except Exception as e:
            self._set_format_designer_status(f"Failed to import format file: {e}")
            QMessageBox.critical(None, "Import Failed", f"Failed to import format file:\n{e}")

    @pyqtSlot()
    def addFormatDefinition(self):
        name = self._unique_format_name("New Format")
        self.format_model.append({"name": name, "columns": self._default_columns("=C{r}*280")})
        self.formatModelChanged.emit()
        self._refresh_xml_type_options()
        self._autosave_format_model_changes()

    @pyqtSlot(int, result=int)
    def duplicateFormatDefinition(self, index):
        if index < 0 or index >= len(self.format_model):
            return -1
        source = self.format_model[index]
        duplicate_name = self._unique_format_name(f"{source.get('name', 'Format')} Copy")
        duplicate_columns = copy.deepcopy(source.get("columns", self._default_columns("=C{r}*280")))
        self.format_model.append({
            "name": duplicate_name,
            "columns": duplicate_columns
        })
        self.formatModelChanged.emit()
        self._refresh_xml_type_options()
        new_index = len(self.format_model) - 1
        try:
            os.makedirs(self.formats_dir, exist_ok=True)
            target_path = self._resolve_unique_format_file_path(duplicate_name)
            with open(target_path, "w", encoding="utf-8") as fp:
                json.dump(self.format_model[new_index], fp, indent=2)
            self._autosave_formats()
            self._set_format_designer_status(f"Duplicated format and saved file: {target_path}")
        except Exception as e:
            self._set_format_designer_status(f"Duplicated format but failed to save file: {e}")
        return new_index

    @pyqtSlot(int)
    def openFormatForEdit(self, index):
        if index < 0 or index >= len(self.format_model):
            return
        self.beginFormatEdit(index)
        if self.root:
            self.root.setProperty("formatDesignerSelectedFormatIndex", index)
            self.root.setProperty("formatDesignerSelectedRowIndex", -1)
            self.root.setProperty("processState", "formatCreate")

    @pyqtSlot(int)
    def duplicateFormatAndOpen(self, index):
        if index < 0 or index >= len(self.format_model):
            return
        self._format_edit_snapshot = copy.deepcopy(self.format_model)
        self._format_edit_active = True

        source = self.format_model[index]
        duplicate_name = self._unique_format_name(f"{source.get('name', 'Format')} Copy")
        duplicate_columns = copy.deepcopy(source.get("columns", self._default_columns("=C{r}*280")))
        self.format_model.append({
            "name": duplicate_name,
            "columns": duplicate_columns
        })
        self.formatModelChanged.emit()
        self._refresh_xml_type_options()
        self._autosave_format_model_changes()

        new_index = len(self.format_model) - 1
        if self.root:
            self.root.setProperty("formatDesignerSelectedFormatIndex", new_index)
            self.root.setProperty("formatDesignerSelectedRowIndex", -1)
            self.root.setProperty("processState", "formatCreate")

    @pyqtSlot(result=int)
    def createFormatDraft(self):
                                                                                 
        self._format_edit_snapshot = copy.deepcopy(self.format_model)
        self._format_edit_active = True
        name = self._unique_format_name("New Format")
        self.format_model.append({"name": name, "columns": self._default_columns("=C{r}*280")})
        self.formatModelChanged.emit()
        self._refresh_xml_type_options()
        self._autosave_format_model_changes()
        return len(self.format_model) - 1

    @pyqtSlot(int)
    def beginFormatEdit(self, format_index):
        if format_index < 0 or format_index >= len(self.format_model):
            return
        self._format_edit_snapshot = copy.deepcopy(self.format_model)
        self._format_edit_active = True

    @pyqtSlot()
    def cancelFormatEdit(self):
        if not self._format_edit_active or self._format_edit_snapshot is None:
            return
        self.format_model = copy.deepcopy(self._format_edit_snapshot)
        self._format_edit_snapshot = None
        self._format_edit_active = False
        self.formatModelChanged.emit()
        self._refresh_xml_type_options()
        self._autosave_format_model_changes()
        self._set_format_designer_status("Discarded unsaved format changes.")

    @pyqtSlot(result=bool)
    def confirmDiscardFormatEdit(self):
        if not self._format_edit_active:
            return True
        result = QMessageBox.question(
            None,
            "Discard Changes",
            "Unsaved format changes will not be saved.\n\nGo back to list anyway?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        return result == QMessageBox.StandardButton.Yes

    @pyqtSlot()
    def commitFormatEdit(self):
        self._format_edit_snapshot = None
        self._format_edit_active = False
        self._autosave_format_model_changes()

    @pyqtSlot(int)
    def deleteFormatDefinition(self, index):
        if index < 0 or index >= len(self.format_model):
            return
        fmt = self.format_model[index]
        format_name = fmt.get("name", "")
        if self._is_builtin_format_name(format_name):
            self._set_format_designer_status("Failed to delete: Den, Glacier, and Globe are built-in formats.")
            return
        alias_names = []
        raw_aliases = fmt.get("__aliases", [])
        if isinstance(raw_aliases, list):
            alias_names = [str(v).strip() for v in raw_aliases if str(v).strip()]
        names_to_remove = [format_name] + alias_names
        deleted_paths = []
        try:
            os.makedirs(self.formats_dir, exist_ok=True)
            deleted_paths = self._delete_format_files_for_names(names_to_remove)
        except Exception as e:
            self._set_format_designer_status(f"Failed to delete format file: {e}")
            return
        self.format_model.pop(index)
        if not self.format_model:
            self.format_model = self._default_formats()
        self.formatModelChanged.emit()
        self._refresh_xml_type_options()
        self._persist_formats_after_delete()
        if deleted_paths:
            self._set_format_designer_status(f"Deleted format and removed {len(deleted_paths)} file(s).")
        else:
            self._set_format_designer_status("Deleted format from model. No matching sidecar format file found.")

    @pyqtSlot(int, str)
    def renameFormatDefinition(self, index, name):
        if index < 0 or index >= len(self.format_model):
            return
        current_name = str(self.format_model[index].get("name", "")).strip()
        next_name = self._unique_format_name(name, skip_index=index)
        if current_name and current_name.lower() != next_name.lower():
            aliases = self.format_model[index].get("__aliases", [])
            if not isinstance(aliases, list):
                aliases = []
            lowered = {str(v).strip().lower() for v in aliases}
            if current_name.lower() not in lowered:
                aliases.append(current_name)
            self.format_model[index]["__aliases"] = aliases
        self.format_model[index]["name"] = next_name
        self.formatModelChanged.emit()
        self._refresh_xml_type_options()
        self._autosave_format_model_changes()

    @pyqtSlot(int, result=int)
    def addFormatRow(self, format_index):
        if format_index < 0 or format_index >= len(self.format_model):
            return -1
        columns = self.format_model[format_index]["columns"]
        new_row = {
            "col": self._next_column_label(columns),
            "type": "data",
            "value": "",
            "width": 14,
            "labelKey": "",
        }
        columns.append(new_row)
        self._sort_format_columns(columns)
        new_index = -1
        for i, candidate in enumerate(columns):
            if candidate is new_row:
                new_index = i
                break
        self._autosave_format_model_changes()
        QTimer.singleShot(0, self.formatModelChanged.emit)
        return new_index

    @pyqtSlot(int, int)
    def deleteFormatRow(self, format_index, row_index):
        if format_index < 0 or format_index >= len(self.format_model):
            return
        columns = self.format_model[format_index]["columns"]
        if row_index < 0 or row_index >= len(columns):
            return
        columns.pop(row_index)
        self._autosave_format_model_changes()
        self.formatModelChanged.emit()

    @pyqtSlot(int, int, str, 'QVariant', result=int)
    def updateFormatRow(self, format_index, row_index, field, value):
        if format_index < 0 or format_index >= len(self.format_model):
            return -1
        columns = self.format_model[format_index]["columns"]
        if row_index < 0 or row_index >= len(columns):
            return -1
        row = columns[row_index]
        if field == "col":
            normalized = self._sanitize_column_input(value)
            if normalized:
                row["col"] = normalized
                self._sort_format_columns(columns)
        elif field == "type":
            row["type"] = self._sanitize_format_type(value)
            row["value"] = self._sanitize_format_value(row["type"], row.get("value", ""))
            row["labelKey"] = self._sanitize_label_key(row.get("labelKey", ""), row["type"])
        elif field == "value":
            row["value"] = self._sanitize_format_value(row.get("type", "data"), value)
        elif field == "width":
            row["width"] = self._sanitize_format_width(value)
        elif field == "labelKey":
            row["labelKey"] = self._sanitize_label_key(value, row.get("type", "data"))
        updated_index = -1
        for i, candidate in enumerate(columns):
            if candidate is row:
                updated_index = i
                break
        if updated_index < 0:
            updated_index = row_index
        self._autosave_format_model_changes()
        QTimer.singleShot(0, self.formatModelChanged.emit)
        return updated_index

    @pyqtSlot(int, int, int, result=int)
    def moveFormatRow(self, format_index, from_index, to_index):
        if format_index < 0 or format_index >= len(self.format_model):
            return -1
        columns = self.format_model[format_index]["columns"]
        if not isinstance(columns, list) or not columns:
            return -1
        if from_index < 0 or from_index >= len(columns):
            return -1
        safe_to = max(0, min(int(to_index), len(columns) - 1))
        if from_index == safe_to:
            return from_index
        moved = columns.pop(from_index)
        columns.insert(safe_to, moved)
        self._autosave_format_model_changes()
        QTimer.singleShot(0, self.formatModelChanged.emit)
        return safe_to

    @pyqtSlot(int, result=bool)
    def loadXmlPreview(self, max_rows=10):
        source_file = self._preview_source_file()
        if not source_file:
            self._set_xml_preview([], [], "Select an XML file first.")
            return False
        rows_limit = max(1, min(30, int(max_rows) if max_rows else 10))
        try:
            df_xml = pd.read_xml(
                source_file,
                xpath=".//ns:Items",
                namespaces={"ns": "http://tempuri.org/ArrayFieldDataSet.xsd"}
            )
            if df_xml is None or df_xml.empty:
                self._set_xml_preview([], [], f"No rows found in {os.path.basename(source_file)}.")
                return False
            df_filtered = df_xml[~df_xml.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)
            if df_filtered.empty:
                self._set_xml_preview([], [], f"No valid rows found in {os.path.basename(source_file)}.")
                return False
            df_data = df_filtered.iloc[:, 1:] if df_filtered.shape[1] > 1 else df_filtered
            max_cols = min(12, df_data.shape[1])
            if max_cols <= 0:
                self._set_xml_preview([], [], f"No previewable columns in {os.path.basename(source_file)}.")
                return False
            headers = []
            for i in range(max_cols):
                headers.append({
                    "index": i,
                    "name": str(df_data.columns[i])
                })
            rows = []
            for r in range(min(rows_limit, len(df_data))):
                row_vals = []
                for c in range(max_cols):
                    val = df_data.iloc[r, c]
                    row_vals.append("" if pd.isna(val) else str(val))
                rows.append(row_vals)
            self._set_xml_preview(headers, rows, f"Previewing {os.path.basename(source_file)}")
            return True
        except Exception as e:
            self._set_xml_preview([], [], f"Preview failed: {e}")
            return False

    @pyqtSlot()
    def selectPreviewXmlFile(self):
        if self.preview_selected_file and os.path.exists(self.preview_selected_file):
            self.loadXmlPreview(10)
            return
        self.selectAnotherPreviewXmlFile()

    @pyqtSlot()
    def selectAnotherPreviewXmlFile(self):
        start_dir = self.last_open_dir if self.last_open_dir and os.path.isdir(self.last_open_dir) else ""
        file_path, _ = QFileDialog.getOpenFileName(None, "Select XML Preview File", start_dir, "XML Files (*.xml)")
        if not file_path:
            return
        self.rememberOpenDirectory(file_path)
        self.preview_selected_file = file_path
        self.loadXmlPreview(10)

    @pyqtSlot(int, int, int, result=int)
    def setFormatRowFromPreview(self, format_index, row_index, column_index):
        if format_index < 0 or format_index >= len(self.format_model):
            return -1
        columns = self.format_model[format_index]["columns"]
        if row_index < 0 or row_index >= len(columns):
            return -1
        safe_col_index = max(0, int(column_index))
        row = columns[row_index]
        row["type"] = "data"
        row["value"] = str(safe_col_index)
        updated_index = -1
        for i, candidate in enumerate(columns):
            if candidate is row:
                updated_index = i
                break
        if updated_index < 0:
            updated_index = row_index
        self._autosave_format_model_changes()
        QTimer.singleShot(0, self.formatModelChanged.emit)
        return updated_index

    @pyqtSlot()
    def saveFormatModel(self):
        try:
            target_path = self.formats_path
            os.makedirs(self.formats_dir, exist_ok=True)
            self._prepare_file_for_write(target_path)
            with open(target_path, "w", encoding="utf-8") as fp:
                json.dump(self.format_model, fp, indent=2)
            self._hide_file_if_supported(target_path)
            self._refresh_xml_type_options()
            self._set_format_designer_status(f"Saved formats to {target_path}")
            QMessageBox.information(None, "Formats Saved", f"Formats saved to:\n{target_path}")
        except Exception as e:
            self._set_format_designer_status(f"Failed to save format: {e}")

    @pyqtSlot(int)
    def saveFormatByName(self, format_index):
        if format_index < 0 or format_index >= len(self.format_model):
            return
        try:
            os.makedirs(self.formats_dir, exist_ok=True)
            fmt = self.format_model[format_index]
            if self._is_builtin_format_name(fmt.get("name", "")):
                self._set_format_designer_status("Built-in formats are not saved as individual files.")
                return
            file_name = self._safe_format_filename(fmt.get("name", "format"))
            target_path = self._resolve_unique_format_file_path(fmt.get("name", "format"))
            with open(target_path, "w", encoding="utf-8") as fp:
                json.dump(fmt, fp, indent=2)
            self._autosave_formats()
            self._set_format_designer_status(f"Saved format to {target_path}")
        except Exception as e:
            self._set_format_designer_status(f"Failed to save format: {e}")

    def rememberOpenDirectory(self, file_path):
        directory = os.path.dirname(file_path) if file_path else ""
        if not directory:
            return
        self.last_open_dir = directory
        self.settings.setValue("lastOpenDir", directory)

    def rememberSaveDirectory(self, directory, batch=False):
        if not directory:
            return
        self.last_save_dir = directory
        self.settings.setValue("lastSaveDir", directory)
        if batch:
            self.last_batch_dir = directory
            self.settings.setValue("lastBatchDir", directory)

    def applySelectedPaths(self, file_paths):
        if not file_paths:
            QMessageBox.information(None, "Info", "No XML files selected")
            return

        is_batch_selection = len(file_paths) > 1
        self.selected_files = file_paths if is_batch_selection else []
        self.selected_file = file_paths[0] if not is_batch_selection else None
        self.is_batch = is_batch_selection
        self.current_batch_index = 0
        self.xml_type = ""
        self.batch_file_statuses = ["Queued"] * len(self.selected_files) if is_batch_selection else []
        self.preview_selected_file = ""
        self._set_xml_preview([], [], "")

        if self.root:
                                                                                       
            self.root.setProperty("processState", "")
            self.root.setProperty("isBatch", False)
            self.root.setProperty("selectedFiles", [])
            self.root.setProperty("batchFileStatuses", [])
            self.root.setProperty("selectedFile", "")
            self.root.setProperty("totalBatchFiles", 0)
            self.root.setProperty("currentBatchIndex", 0)
            self.root.setProperty("currentFileName", "")
            QApplication.processEvents()
            self.root.setProperty("selectionType", "")
            self.root.setProperty("selectedFile", self.selected_file or "")
            self.root.setProperty("selectedFiles", list(self.selected_files) if is_batch_selection else [])
            self.root.setProperty("isBatch", is_batch_selection)
            self.root.setProperty("totalBatchFiles", len(file_paths) if is_batch_selection else 0)
            self.root.setProperty("currentBatchIndex", 0)
            self.root.setProperty("currentFileName", os.path.basename(file_paths[0]) if is_batch_selection else "")
            self.root.setProperty("batchOutputs", [])
            self.refreshBatchFileStatusesProperty()
            if is_batch_selection:
                self.root.setProperty("fileSize", "")
            else:
                self.root.setProperty("fileSize", self.getFileSize(self.selected_file))
            self.root.setProperty("processState", "selecting")

    @pyqtSlot()
    def selectFile(self):
        start_dir = self.last_open_dir if self.last_open_dir and os.path.isdir(self.last_open_dir) else ""
        file_paths, _ = QFileDialog.getOpenFileNames(None, "Select XML File(s)", start_dir, "XML Files (*.xml)")
        if file_paths:
            self.rememberOpenDirectory(file_paths[0])
        self.applySelectedPaths(file_paths)

    @pyqtSlot('QVariantList')
    def setDroppedPaths(self, paths):
        self.cancel_requested = False
        normalized_paths = [normalize_path(p) for p in paths if normalize_path(p)]
        if not normalized_paths:
            QMessageBox.warning(None, "Invalid Selection", "No valid dropped paths were found.")
            return
        if self.path_scan_thread and self.path_scan_thread.isRunning():
            QMessageBox.information(None, "Scanning", "Please wait for the current folder scan to finish.")
            return

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        self._stop_thread("path_scan_thread", "path_scan_worker")
        self.path_scan_thread = QThread()
        self.path_scan_worker = PathDiscoveryWorker(normalized_paths)
        self.path_scan_worker.moveToThread(self.path_scan_thread)
        self.path_scan_thread.started.connect(self.path_scan_worker.process)
        self.path_scan_worker.finished.connect(self.handleDroppedPathScanFinished)
        self.path_scan_worker.error.connect(self.handleDroppedPathScanError)
        self.path_scan_worker.finished.connect(self.path_scan_thread.quit)
        self.path_scan_worker.error.connect(self.path_scan_thread.quit)
        self.path_scan_thread.finished.connect(self.cleanupDroppedPathScan)
        self.path_scan_thread.start()

    @pyqtSlot(object)
    def handleDroppedPathScanFinished(self, xml_paths):
        if self.cancel_requested:
            return
        if xml_paths:
            self.applySelectedPaths(xml_paths)
            return
        QMessageBox.warning(None, "Invalid Selection", "No XML files were found in the dropped item(s).")

    @pyqtSlot(str)
    def handleDroppedPathScanError(self, msg):
        if msg == "Operation cancelled by user.":
            return
        QMessageBox.critical(None, "Error", msg)

    def cleanupDroppedPathScan(self):
        if QApplication.overrideCursor() is not None:
            QApplication.restoreOverrideCursor()
        self._stop_thread("path_scan_thread", "path_scan_worker")


    @pyqtSlot()
    def selectBatchFiles(self):
        self.selectFile()

    @pyqtSlot()
    def confirmAndConvertBatch(self):
        if not self.xml_type or not self.selected_files:
            return
        self.cancel_requested = False
        self.is_batch = True
        self.selected_file = None
        self.batch_results = []
        self.batch_outputs = []
        self.batch_file_statuses = ["Queued"] * len(self.selected_files)
        if self.root:
                                                                                         
            self.root.setProperty("processState", "")
            self.root.setProperty("isBatch", False)
            self.root.setProperty("selectedFiles", [])
            self.root.setProperty("selectedFile", "")
            self.root.setProperty("currentBatchIndex", -1)
            self.root.setProperty("totalBatchFiles", 0)
            self.root.setProperty("currentFileName", "")
            self.root.setProperty("batchFileStatuses", [])
            self.root.setProperty("batchOutputs", [])
            QApplication.processEvents()

            self.root.setProperty("processState", "selecting")
            QApplication.processEvents()
            self.root.setProperty("isBatch", True)
            self.root.setProperty("selectedFiles", list(self.selected_files))
            self.root.setProperty("processState", "converting")
            self.root.setProperty("currentBatchIndex", 0)
            self.root.setProperty("totalBatchFiles", len(self.selected_files))
            self.root.setProperty("currentFileName", os.path.basename(self.selected_files[0]))
            self.refreshBatchFileStatusesProperty()
        self.progress = 0
        self.current_batch_index = 0
        self.progressUpdated.emit(self.progress)
        self.processNextBatchFile()

    def processNextBatchFile(self):
        if self.current_batch_index >= len(self.selected_files):
            return
        self.is_batch = len(self.selected_files) > 1
        if self.current_batch_index < len(self.batch_file_statuses):
            self.batch_file_statuses[self.current_batch_index] = "Processing"
            self.refreshBatchFileStatusesProperty()
        if self.root:
            self.root.setProperty("currentBatchIndex", self.current_batch_index)
            self.root.setProperty("totalBatchFiles", len(self.selected_files))
            self.root.setProperty("currentFileName", os.path.basename(self.selected_files[self.current_batch_index]))
        self._stop_thread("thread", "worker")
        self.thread = QThread()
        selected_format = next((fmt for fmt in self.format_model if fmt.get("name", "") == self.xml_type), None)
        self.worker = Worker([self.selected_files[self.current_batch_index]], self.xml_type, selected_format)
        self.worker.moveToThread(self.thread)
        self.worker.progress.connect(self.progressUpdated)
        self.worker.error.connect(self.handleError)
        if self.is_batch:
            self.worker.dataReady.connect(self.collectBatchResult)
        else:
            self.worker.dataReady.connect(self.saveFile)
        self.thread.started.connect(self.worker.process)
        self.thread.start()

    @pyqtSlot(str)
    def setSelectionType(self,type_str):
        self.xml_type = type_str
        self.settings.setValue("lastXmlType", type_str)
        if self.root:
            self.root.setProperty("selectionType",type_str)

    @pyqtSlot('QVariantList', str, bool)
    def syncSelectionContext(self, files, selected_file, is_batch):
        normalized_files = [normalize_path(p) for p in files if normalize_path(p)]
        normalized_selected_file = normalize_path(selected_file) if selected_file else ""
                                                                                   
        if len(normalized_files) > 1:
            self.selected_files = normalized_files
            self.selected_file = None
            self.is_batch = True
            return
        if len(normalized_files) == 1:
            self.selected_files = []
            self.selected_file = normalized_files[0]
            self.is_batch = False
            return
        if normalized_selected_file:
            self.selected_files = []
            self.selected_file = normalized_selected_file
            self.is_batch = False
            return
        if not self.selected_files and not self.selected_file:
            self.is_batch = bool(is_batch)

    def _qml_list(self, value):
        if isinstance(value, (list, tuple)):
            return list(value)
        if hasattr(value, "toVariant"):
            try:
                variant = value.toVariant()
                if isinstance(variant, (list, tuple)):
                    return list(variant)
            except Exception:
                pass
        return []

    @pyqtSlot()
    def confirmAndConvert(self):
        self.cancel_requested = False
                                                                                        
        if self.root:
            raw_files = self.root.property("selectedFiles")
            ui_files = [normalize_path(p) for p in self._qml_list(raw_files) if normalize_path(p)]
            raw_selected = self.root.property("selectedFile")
            ui_selected = normalize_path(raw_selected) if isinstance(raw_selected, str) and raw_selected else ""

            if len(ui_files) > 1:
                self.selected_files = ui_files
                self.selected_file = None
                self.is_batch = True
            elif len(ui_files) == 1:
                self.selected_files = []
                self.selected_file = ui_files[0]
                self.is_batch = False
            elif ui_selected:
                self.selected_files = []
                self.selected_file = ui_selected
                self.is_batch = False
                                                                                                                  

            ui_type = self.root.property("selectionType")
            if ui_type:
                self.xml_type = ui_type

                                                                                          
        if len(self.selected_files) > 1:
            self.is_batch = True
            self.selected_file = None
        elif self.selected_file:
            self.is_batch = False

        if not self.xml_type:
            QMessageBox.information(None, "Info", "Please select XML type before converting.")
            return

        if len(self.selected_files) > 1:
            self.confirmAndConvertBatch()
            return

        if len(self.selected_files) == 1 and not self.selected_file:
            self.selected_file = self.selected_files[0]
            self.selected_files = []
            self.is_batch = False

        if not self.selected_file:
            return

        if self.root:
            self.root.setProperty("processState","converting")
        self.progress=0
        self.progressUpdated.emit(self.progress)
        self._stop_thread("thread", "worker")
        self.thread=QThread()
        selected_format = next((fmt for fmt in self.format_model if fmt.get("name", "") == self.xml_type), None)
        self.worker = Worker([self.selected_file], self.xml_type, selected_format)
        self.worker.moveToThread(self.thread)
        self.worker.progress.connect(self.progressUpdated)
        self.worker.error.connect(self.handleError)
        self.worker.dataReady.connect(self.saveFile)
        self.thread.started.connect(self.worker.process)
        self.thread.start()

    @pyqtSlot()
    def selectDifferentFile(self):
        self.resetProperties()

    @pyqtSlot(str,result=str)
    def getFileSize(self,file_path):
        try:
            size_bytes=os.path.getsize(file_path)
            return f"{size_bytes/1024:.2f} KB" if size_bytes<1024*1024 else f"{size_bytes/(1024*1024):.2f} MB"
        except OSError:
            return "Unknown"

    @pyqtSlot(int)
    def removeSelectedFile(self, index):
        if not self.is_batch:
            return
        if index < 0 or index >= len(self.selected_files):
            return

        self.selected_files.pop(index)
        if index < len(self.batch_file_statuses):
            self.batch_file_statuses.pop(index)

        if not self.selected_files:
            self.resetProperties()
            return

        if len(self.selected_files) == 1:
            self.selected_file = self.selected_files[0]
            self.selected_files = []
            self.is_batch = False
            self.current_batch_index = 0
            if self.root:
                self.root.setProperty("isBatch", False)
                self.root.setProperty("selectedFiles", [])
                self.root.setProperty("selectedFile", self.selected_file)
                self.root.setProperty("totalBatchFiles", 0)
                self.root.setProperty("currentBatchIndex", 0)
                self.root.setProperty("currentFileName", "")
                self.root.setProperty("fileSize", self.getFileSize(self.selected_file))
                self.root.setProperty("batchFileStatuses", [])
            return

        if self.current_batch_index >= len(self.selected_files):
            self.current_batch_index = len(self.selected_files) - 1
        if self.root:
            self.root.setProperty("isBatch", True)
            self.root.setProperty("selectedFile", "")
            self.root.setProperty("selectedFiles", self.selected_files)
            self.root.setProperty("totalBatchFiles", len(self.selected_files))
            self.root.setProperty("currentBatchIndex", self.current_batch_index)
            self.root.setProperty("currentFileName", os.path.basename(self.selected_files[self.current_batch_index]))
            self.refreshBatchFileStatusesProperty()

    @pyqtSlot()
    def convertAnotherFile(self):
        self.cancel_requested = True
        self._request_all_worker_cancels()
        self._stop_all_background_threads()
        self.resetProperties()

    @pyqtSlot(str)
    def setSelectedFile(self,file_path):
        self.selected_file = file_path
        self.selected_files = []
        self.batch_file_statuses = []
        self.batch_results = []
        self.batch_outputs = []
        self.is_batch = False
        self.current_batch_index = 0
        self.preview_selected_file = ""
        self._set_xml_preview([], [], "")
        if self.root:
            self.root.setProperty("isBatch", False)
            self.root.setProperty("selectedFiles", [])
            self.root.setProperty("totalBatchFiles", 0)
            self.root.setProperty("currentBatchIndex", 0)
            self.root.setProperty("currentFileName", "")
            self.root.setProperty("batchOutputs", [])
            self.root.setProperty("batchFileStatuses", [])

    def updateProgressInQML(self,value):
        if self.root:
            self.root.setProperty("progress",value)

    def handleError(self,msg):
        if msg == "Operation cancelled by user.":
            self._stop_thread("thread", "worker")
            if self.root:
                if self.is_batch:
                    self.root.setProperty("processState", "batchReview")
                else:
                    self.root.setProperty("processState", "idle")
            return
        if self.is_batch and self.root and self.root.property("processState") == "converting":
            QMessageBox.warning(None, "Batch Item Failed", msg)
            if self.current_batch_index < len(self.batch_file_statuses):
                self.batch_file_statuses[self.current_batch_index] = "Failed"
                self.refreshBatchFileStatusesProperty()
            self._stop_thread("thread", "worker")
            self.current_batch_index += 1
            if self.root:
                self.root.setProperty("currentBatchIndex", self.current_batch_index)
            if self.current_batch_index < len(self.selected_files):
                self.processNextBatchFile()
            else:
                if self.root:
                    self.root.setProperty("processState", "batchReview")
            return

        QMessageBox.critical(None,"Error",msg)
        if self.root:
            self.root.setProperty("processState","idle")
        self._stop_thread("thread", "worker")

    @pyqtSlot(object, str, str)
    def collectBatchResult(self, df, xml_type, xml_file):
        if self.cancel_requested:
            self._stop_thread("thread", "worker")
            if self.root:
                self.root.setProperty("processState", "batchReview")
            return
        default_output_path = build_default_batch_output_path(xml_file)
        default_dir = self.last_batch_dir if self.last_batch_dir and os.path.isdir(self.last_batch_dir) else os.path.dirname(default_output_path)
        self.batch_results.append({
            "df": df.replace([float('inf'), float('-inf')], 0).fillna(0),
            "xml_type": xml_type,
            "xml_file": xml_file
        })
        self.batch_outputs.append({
            "sourceFile": os.path.basename(xml_file),
            "fileName": os.path.basename(default_output_path),
            "saveDir": default_dir,
            "savePath": os.path.join(default_dir, os.path.basename(default_output_path))
        })
        self.refreshBatchOutputsProperty()
        if self.current_batch_index < len(self.batch_file_statuses):
            self.batch_file_statuses[self.current_batch_index] = "Done"
            self.refreshBatchFileStatusesProperty()
        self._stop_thread("thread", "worker")

        self.current_batch_index += 1
        if self.root:
            self.root.setProperty("currentBatchIndex", self.current_batch_index)
        if self.current_batch_index < len(self.selected_files):
            self.processNextBatchFile()
        else:
            if self.root:
                self.root.setProperty("processState", "batchReview")

    def refreshBatchOutputsProperty(self):
        if self.root:
            self.root.setProperty("batchOutputs", self.batch_outputs)

    @pyqtSlot(str, result=str)
    def validateOutputDirectory(self, directory):
        return get_invalid_output_directory_message(directory)

    @pyqtSlot('QVariantList', result='QVariantList')
    def estimateBatchOutputConflicts(self, outputs):
        entries = []
        if isinstance(outputs, list):
            entries = outputs
        elif hasattr(outputs, "toVariant"):
            try:
                variant = outputs.toVariant()
                entries = variant if isinstance(variant, list) else []
            except Exception:
                entries = []

        by_path = {}
        for i, item in enumerate(entries):
            if not isinstance(item, dict):
                continue
            file_name = ensure_xlsx_extension(str(item.get("fileName", "")))
            save_dir = str(item.get("saveDir", ""))
            if not file_name or not save_dir:
                continue
            full_path = normalize_path(os.path.join(save_dir, file_name))
            if not full_path:
                continue
            key = os.path.normcase(full_path)
            by_path.setdefault(key, []).append((i, item, full_path))

        conflicts = []
        for _, group in by_path.items():
            if len(group) <= 1:
                continue
            for idx, item, full_path in group:
                conflicts.append({
                    "index": idx,
                    "sourceFile": str(item.get("sourceFile", f"Item {idx + 1}")),
                    "path": full_path,
                    "reason": "Duplicate output path in batch list."
                })
        return conflicts

    @pyqtSlot(int, str)
    def updateBatchOutputFileName(self, index, file_name):
        if index < 0 or index >= len(self.batch_outputs):
            return
        safe_name = normalize_batch_output_name(file_name)
        self.batch_outputs[index]["fileName"] = safe_name
        self.batch_outputs[index]["savePath"] = os.path.join(self.batch_outputs[index]["saveDir"], safe_name)
        self.refreshBatchOutputsProperty()

    @pyqtSlot(int, str)
    def updateBatchOutputDirectory(self, index, directory):
        if index < 0 or index >= len(self.batch_outputs):
            return
        if not directory:
            return
        self.batch_outputs[index]["saveDir"] = directory
        self.batch_outputs[index]["savePath"] = os.path.join(directory, self.batch_outputs[index]["fileName"])
        self.rememberSaveDirectory(directory, batch=True)
        self.refreshBatchOutputsProperty()

    @pyqtSlot(str)
    def applyBatchOutputDirectoryToAll(self, directory):
        if not directory or not self.batch_outputs:
            return
        for i in range(len(self.batch_outputs)):
            self.batch_outputs[i]["saveDir"] = directory
            self.batch_outputs[i]["savePath"] = os.path.join(directory, self.batch_outputs[i]["fileName"])
        self.rememberSaveDirectory(directory, batch=True)
        self.refreshBatchOutputsProperty()

    @pyqtSlot(int)
    def browseBatchOutputDirectory(self, index):
        if index < 0 or index >= len(self.batch_outputs):
            return
        start_dir = self.batch_outputs[index]["saveDir"] if self.batch_outputs[index]["saveDir"] else self.last_batch_dir
        chosen_dir = QFileDialog.getExistingDirectory(None, "Select Save Folder", start_dir)
        if not chosen_dir:
            return
        self.updateBatchOutputDirectory(index, chosen_dir)

    @pyqtSlot()
    def browseBatchOutputDirectoryForAll(self):
        if not self.batch_outputs:
            return
        start_dir = self.batch_outputs[0]["saveDir"] if self.batch_outputs[0]["saveDir"] else self.last_batch_dir
        chosen_dir = QFileDialog.getExistingDirectory(None, "Select Save Folder for All Files", start_dir)
        if not chosen_dir:
            return
        self.applyBatchOutputDirectoryToAll(chosen_dir)

    @pyqtSlot()
    def saveAllBatchOutputs(self):
        if not self.batch_results or not self.batch_outputs:
            return
        self.cancel_requested = False
        dir_issues = []
        for i, output in enumerate(self.batch_outputs):
            reason = get_invalid_output_directory_message(output.get("saveDir", ""))
            if reason:
                src = output.get("sourceFile", f"Item {i + 1}")
                dir_issues.append(f"{i + 1}. {src}: {reason}")
        if dir_issues:
            suffix = "" if len(dir_issues) <= 8 else f"\n...and {len(dir_issues) - 8} more issue(s)"
            QMessageBox.warning(
                None,
                "Invalid Save Folder",
                "Please fix these save folder issues before confirming:\n\n"
                + "\n".join(dir_issues[:8])
                + suffix
            )
            return
        issues = []
        for i, output in enumerate(self.batch_outputs):
            reason = get_invalid_batch_name_message(output.get("fileName", ""))
            if reason:
                src = output.get("sourceFile", f"Item {i + 1}")
                issues.append(f"{i + 1}. {src}: {reason}")
        if issues:
            suffix = "" if len(issues) <= 8 else f"\n...and {len(issues) - 8} more issue(s)"
            QMessageBox.warning(
                None,
                "Invalid Output File Name",
                "Please fix these file name issues before confirming:\n\n"
                + "\n".join(issues[:8])
                + suffix
            )
            return
        conflicts = self.estimateBatchOutputConflicts(self.batch_outputs)
        if conflicts:
            lines = []
            for i, conflict in enumerate(conflicts[:8]):
                lines.append(
                    f"{i + 1}. {conflict.get('sourceFile', 'Item')} -> {conflict.get('path', '')}"
                )
            suffix = "" if len(conflicts) <= 8 else f"\n...and {len(conflicts) - 8} more conflict(s)"
            QMessageBox.warning(
                None,
                "Conflicting Output Paths",
                "Please resolve duplicate output paths before confirming:\n\n"
                + "\n".join(lines)
                + suffix
            )
            return
        target_paths = [
            os.path.join(output["saveDir"], ensure_xlsx_extension(output["fileName"]))
            for output in self.batch_outputs
        ]
        if not confirm_overwrite_paths(target_paths, "Confirm Batch Overwrite"):
            return
        self._start_batch_save_thread()

    def _start_batch_save_thread(self):
        if self.root:
            self.root.setProperty("processState", "creating")
        self.progressUpdated.emit(0)
        self._stop_thread("batch_save_thread", "batch_save_worker")
        self.batch_save_thread = QThread()
        self.batch_save_worker = BatchSaveWorker(self.batch_results, self.batch_outputs)
        self.batch_save_worker.moveToThread(self.batch_save_thread)
        self.batch_save_worker.progress.connect(self.progressUpdated)
        self.batch_save_worker.error.connect(self.handleBatchSaveError)
        self.batch_save_worker.finished.connect(self.handleBatchSaveFinished)
        self.batch_save_thread.started.connect(self.batch_save_worker.save_all)
        self.batch_save_thread.start()

    def handleBatchSaveError(self, msg):
        if msg == "Operation cancelled by user.":
            if self.root:
                self.root.setProperty("processState", "batchReview")
            self._stop_thread("batch_save_thread", "batch_save_worker")
            return
        if str(msg).startswith("BATCH_PERMISSION_DENIED::"):
            parts = str(msg).split("::", 3)
            locked_path = parts[2] if len(parts) > 2 else ""
            self._stop_thread("batch_save_thread", "batch_save_worker")
            result = QMessageBox.question(
                None,
                "File In Use",
                "Cannot overwrite because the file is currently open or in use:\n\n"
                f"{locked_path}\n\n"
                "Close the file, then click Retry.",
                QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Retry
            )
            if result == QMessageBox.StandardButton.Retry:
                self._start_batch_save_thread()
                return
            if self.root:
                self.root.setProperty("processState", "batchReview")
            return
        QMessageBox.critical(None, "Error", msg)
        if self.root:
            self.root.setProperty("processState", "batchReview")
        self._stop_thread("batch_save_thread", "batch_save_worker")

    @pyqtSlot(object)
    def handleBatchSaveFinished(self, saved_outputs):
        if self.cancel_requested:
            if self.root:
                self.root.setProperty("processState", "batchReview")
            self._stop_thread("batch_save_thread", "batch_save_worker")
            return
        self.batch_outputs = saved_outputs
        self.refreshBatchOutputsProperty()
        self.progressUpdated.emit(100)
        if self.root:
            self.root.setProperty("processState", "complete")
        QMessageBox.information(None, "Done", f"Saved {len(self.batch_outputs)} files successfully.")
        self._stop_thread("batch_save_thread", "batch_save_worker")

    @pyqtSlot(object,str,str)
    def saveFile(self, df, xml_type, xml_file):
        df = df.replace([float('inf'), float('-inf')], 0).fillna(0)
        start_dir = self.last_save_dir if self.last_save_dir and os.path.isdir(self.last_save_dir) else os.path.dirname(xml_file)
        default_name = os.path.splitext(os.path.basename(xml_file))[0] + ".xlsx"
        default_path = os.path.join(start_dir, default_name)
        save_path, _ = QFileDialog.getSaveFileName(
            None,
            "Save Excel File",
            default_path,
            "Excel Files (*.xlsx)"
        )
        if save_path and not save_path.lower().endswith(".xlsx"): save_path += ".xlsx"
        if not save_path:
            self._stop_thread("thread", "worker")
            self.resetProperties()
            return
        self.rememberSaveDirectory(os.path.dirname(save_path), batch=False)
        if not confirm_overwrite_paths([save_path], "Confirm Overwrite"):
            self._stop_thread("thread", "worker")
            self.resetProperties()
            return

        self._last_save_payload = {
            "df": df,
            "xml_type": xml_type,
            "xml_file": xml_file,
            "save_path": save_path,
        }
        self._start_single_save_thread(df, xml_type, save_path, xml_file)
        self._stop_thread("thread", "worker")

    def _start_single_save_thread(self, df, xml_type, save_path, xml_file):
        self.progressUpdated.emit(90)
        self._stop_thread("save_thread", "save_worker")
        self.save_thread = QThread()
        self.save_worker = SaveWorker(df, xml_type, save_path, xml_file)
        self.save_worker.moveToThread(self.save_thread)
        self.save_worker.progress.connect(self.progressUpdated)
        self.save_worker.error.connect(self.handleSaveError)
        self.save_worker.saved.connect(self.handleSaved)
        self.save_thread.started.connect(self.save_worker.save)
        self.save_thread.start()

    def handleSaveError(self, msg):
        if msg == "Operation cancelled by user.":
            if self.root:
                self.root.setProperty("processState", "idle")
            self._stop_thread("save_thread", "save_worker")
            return
        if str(msg).startswith("PERMISSION_DENIED::"):
            parts = str(msg).split("::", 2)
            locked_path = parts[1] if len(parts) > 1 else ""
            self._stop_thread("save_thread", "save_worker")
            result = QMessageBox.question(
                None,
                "File In Use",
                "Cannot overwrite because the file is currently open or in use:\n\n"
                f"{locked_path}\n\n"
                "Close the file, then click Retry.",
                QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Retry
            )
            if result == QMessageBox.StandardButton.Retry and self._last_save_payload:
                payload = self._last_save_payload
                self._start_single_save_thread(
                    payload.get("df"),
                    payload.get("xml_type", ""),
                    payload.get("save_path", ""),
                    payload.get("xml_file", "")
                )
                return
            if self.root:
                self.root.setProperty("processState", "idle")
            return
        QMessageBox.critical(None, "Error", msg)
        if self.root:
            self.root.setProperty("processState", "idle")
        self._stop_thread("save_thread", "save_worker")

    @pyqtSlot(str, str)
    def handleSaved(self, save_path, xml_file):
        if self.cancel_requested:
            if self.root:
                self.root.setProperty("processState", "idle")
            self._stop_thread("save_thread", "save_worker")
            return
        self.progressUpdated.emit(100)
        QMessageBox.information(None, "Done", f"Processed Excel saved:\n{save_path}")
        self.current_batch_index += 1
        if self.root:
            self.root.setProperty("currentBatchIndex", self.current_batch_index)
        if self.current_batch_index < len(self.selected_files):
            self.processNextBatchFile()
        else:
            if self.root:
                self.root.setProperty("processState", "complete")
            self._last_save_payload = None
        self._stop_thread("save_thread", "save_worker")

    @pyqtSlot()
    def cancelCurrentOperation(self):
        self.cancel_requested = True
        self._request_all_worker_cancels()
        self._stop_all_background_threads()

        if self.is_batch:
            for i, status in enumerate(self.batch_file_statuses):
                if status in ("Queued", "Processing"):
                    self.batch_file_statuses[i] = "Cancelled"
            self.refreshBatchFileStatusesProperty()
            if self.root:
                self.root.setProperty("processState", "batchReview")
            return

        if self.root:
            self.root.setProperty("processState", "idle")

    def resetProperties(self):
        self._stop_all_background_threads()
        self.preview_selected_file = ""
        self._set_xml_preview([], [], "")
        if self.root:
            self.root.setProperty("processState","idle")
            self.root.setProperty("selectedFile","")
            self.root.setProperty("selectedFiles", [])
            self.root.setProperty("selectionType","")
            self.root.setProperty("fileSize","")
            self.root.setProperty("progress",0)
            self.root.setProperty("isBatch", False)
            self.root.setProperty("totalBatchFiles", 0)
            self.root.setProperty("currentBatchIndex", 0)
            self.root.setProperty("currentFileName", "")
            self.root.setProperty("batchOutputs", [])
            self.root.setProperty("batchFileStatuses", [])
        self.selected_file = None
        self.selected_files = []
        self.batch_file_statuses = []
        self.batch_results = []
        self.batch_outputs = []
        self.is_batch = False
        self.current_batch_index = 0
        self.xml_type=""
        self.progress=0
        self.save_thread = None
        self.save_worker = None
        self.batch_save_thread = None
        self.batch_save_worker = None

if __name__=="__main__":
    app=QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    app.setStyle("Fusion")
    app_font_family = app.font().family()
    font_path = os.path.join(RESOURCE_BASE_DIR, "fonts", "Minecraft.ttf")
    if os.path.exists(font_path):
        font_id = QFontDatabase.addApplicationFont(font_path)
        if font_id != -1:
            loaded_families = QFontDatabase.applicationFontFamilies(font_id)
            if loaded_families:
                app_font_family = loaded_families[0]
                app.setFont(QFont(app_font_family))
    icon_path=os.path.join(RESOURCE_BASE_DIR,"images","icon.png")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    engine=QQmlApplicationEngine()

    backend=Backend(engine)
    engine.rootContext().setContextProperty("backend",backend)
    engine.rootContext().setContextProperty("appFontFamily", app_font_family)

    qml_file=os.path.join(RESOURCE_BASE_DIR,"main.qml")
    engine.load(QUrl.fromLocalFile(qml_file))
    if not engine.rootObjects(): sys.exit(-1)

    backend.root=engine.rootObjects()[0]
    if os.path.exists(icon_path):
        backend.root.setIcon(QIcon(icon_path))
    backend.applyRememberedSettingsToUI()
    sys.exit(app.exec())
