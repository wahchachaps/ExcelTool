import sys
import os
import pandas as pd
import win32com.client as win32
from PyQt6.QtCore import QObject, pyqtSlot, QUrl, pyqtSignal, QThread, Qt, QSettings
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox
from PyQt6.QtQml import QQmlApplicationEngine


def ensure_xlsx_extension(file_name):
    if not file_name:
        return "output.xlsx"
    return file_name if file_name.lower().endswith(".xlsx") else f"{file_name}.xlsx"


def build_default_batch_output_path(xml_file):
    source_dir = os.path.dirname(xml_file)
    source_base = os.path.splitext(os.path.basename(xml_file))[0]
    return os.path.join(source_dir, f"{source_base}_converted.xlsx")


def normalize_path(path):
    if not path:
        return ""
    normalized = os.path.normpath(str(path))
    return normalized


def collect_xml_files_from_paths(paths):
    collected = []
    seen = set()
    for raw_path in paths:
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
                for file_name in files:
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
    with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
        xml_file_name = os.path.splitext(os.path.basename(xml_file))[0]
        sheet_name = ('_'.join(xml_file_name.split('_')[:-1]) if '_' in xml_file_name else xml_file_name)[:31]
        df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
        workbook = writer.book
        ws = writer.sheets[sheet_name]

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

    def __init__(self, xml_files, xml_type):
        super().__init__()
        self.xml_files = xml_files
        self.xml_type = xml_type
        self.current_file_index = 0

    @pyqtSlot()
    def process(self):
        if self.current_file_index >= len(self.xml_files):
            return
        xml_file = self.xml_files[self.current_file_index]
        try:
            if self.xml_type not in ["Den", "Globe", "Glacier"]:
                self.error.emit("Only Den, Glacier, and Globe are implemented.")
                return
            try:
                df_xml = pd.read_xml(
                    xml_file,
                    xpath=".//ns:Items",
                    namespaces={"ns": "http://tempuri.org/ArrayFieldDataSet.xsd"}
                )
            except Exception as e:
                self.error.emit(f"Error reading XML {xml_file}: {e}")
                self.current_file_index += 1
                self.process()
                return
            self.progress.emit(50)
            if self.xml_type == "Den": self.process_den_from_df(df_xml, xml_file)
            elif self.xml_type == "Globe": self.process_globe_from_df(df_xml, xml_file)
            elif self.xml_type == "Glacier": self.process_glacier_from_df(df_xml, xml_file)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred for {xml_file}: {e}")
            self.current_file_index += 1
            self.process()

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

    @pyqtSlot()
    def save(self):
        try:
            self.progress.emit(90)
            export_dataframe_to_excel(self.df, self.xml_type, self.save_path, self.xml_file)
            self.progress.emit(100)
            self.saved.emit(self.save_path, self.xml_file)
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

    @pyqtSlot()
    def save_all(self):
        try:
            total = len(self.batch_results)
            if total == 0:
                self.finished.emit(self.batch_outputs)
                return

            for i, result in enumerate(self.batch_results):
                output = self.batch_outputs[i]
                save_path = os.path.join(output["saveDir"], ensure_xlsx_extension(output["fileName"]))
                export_dataframe_to_excel(result["df"], result["xml_type"], save_path, result["xml_file"])
                self.batch_outputs[i]["savePath"] = save_path
                self.batch_outputs[i]["fileName"] = os.path.basename(save_path)
                self.progress.emit(int(((i + 1) / total) * 100))

            self.finished.emit(self.batch_outputs)
        except Exception as e:
            self.error.emit(f"Failed to save batch files: {e}")


class Backend(QObject):
    progressUpdated = pyqtSignal(int)

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
        self.batch_results = []
        self.batch_outputs = []
        self.batch_file_statuses = []
        self.settings = QSettings("ExcelTool", "ExcelTool")
        self.last_open_dir = str(self.settings.value("lastOpenDir", "", str))
        self.last_save_dir = str(self.settings.value("lastSaveDir", "", str))
        self.last_batch_dir = str(self.settings.value("lastBatchDir", "", str))
        self.xml_type = str(self.settings.value("lastXmlType", "", str))

    def refreshBatchFileStatusesProperty(self):
        if self.root:
            self.root.setProperty("batchFileStatuses", self.batch_file_statuses)

    def applyRememberedSettingsToUI(self):
        if not self.root:
            return
        if self.xml_type:
            self.root.setProperty("selectionType", self.xml_type)

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

        if self.root:
            # Force change notifications even when selecting the same files repeatedly.
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
        xml_paths = collect_xml_files_from_paths(paths)
        if not xml_paths:
            QMessageBox.warning(None, "Invalid Selection", "No XML files were found in the dropped item(s).")
            return
        self.applySelectedPaths(xml_paths)


    @pyqtSlot()
    def selectBatchFiles(self):
        self.selectFile()

    @pyqtSlot()
    def confirmAndConvertBatch(self):
        if not self.xml_type or not self.selected_files:
            return
        self.is_batch = True
        self.selected_file = None
        self.batch_results = []
        self.batch_outputs = []
        self.batch_file_statuses = ["Queued"] * len(self.selected_files)
        if self.root:
            # Force a full UI refresh so second/third runs with same files still repaint.
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
        self.thread = QThread()
        self.worker = Worker([self.selected_files[self.current_batch_index]], self.xml_type)
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
        # Do not clobber a valid backend selection with an empty/stale QML payload.
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
        # Always re-sync from QML so repeated runs with same files are handled reliably.
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
            # If QML reports empty selection, keep backend selection from applySelectedPaths/syncSelectionContext.

            ui_type = self.root.property("selectionType")
            if ui_type:
                self.xml_type = ui_type

        # Derive mode from current selection to avoid stale is_batch across repeated runs.
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
        self.thread=QThread()
        self.worker = Worker([self.selected_file], self.xml_type)
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
        if self.is_batch and self.root and self.root.property("processState") == "converting":
            QMessageBox.warning(None, "Batch Item Failed", msg)
            if self.current_batch_index < len(self.batch_file_statuses):
                self.batch_file_statuses[self.current_batch_index] = "Failed"
                self.refreshBatchFileStatusesProperty()
            if self.thread:
                self.thread.quit()
                self.thread.wait()
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
        if self.thread:
            self.thread.quit()
            self.thread.wait()

    @pyqtSlot(object, str, str)
    def collectBatchResult(self, df, xml_type, xml_file):
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
        if self.thread:
            self.thread.quit()
            self.thread.wait()

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

    @pyqtSlot(int, str)
    def updateBatchOutputFileName(self, index, file_name):
        if index < 0 or index >= len(self.batch_outputs):
            return
        safe_name = ensure_xlsx_extension(file_name.strip())
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
        target_paths = [
            os.path.join(output["saveDir"], ensure_xlsx_extension(output["fileName"]))
            for output in self.batch_outputs
        ]
        if not confirm_overwrite_paths(target_paths, "Confirm Batch Overwrite"):
            return
        if self.root:
            self.root.setProperty("processState", "creating")
        self.progressUpdated.emit(0)
        self.batch_save_thread = QThread()
        self.batch_save_worker = BatchSaveWorker(self.batch_results, self.batch_outputs)
        self.batch_save_worker.moveToThread(self.batch_save_thread)
        self.batch_save_worker.progress.connect(self.progressUpdated)
        self.batch_save_worker.error.connect(self.handleBatchSaveError)
        self.batch_save_worker.finished.connect(self.handleBatchSaveFinished)
        self.batch_save_thread.started.connect(self.batch_save_worker.save_all)
        self.batch_save_thread.start()

    def handleBatchSaveError(self, msg):
        QMessageBox.critical(None, "Error", msg)
        if self.root:
            self.root.setProperty("processState", "batchReview")
        if self.batch_save_thread:
            self.batch_save_thread.quit()
            self.batch_save_thread.wait()
        self.batch_save_thread = None
        self.batch_save_worker = None

    @pyqtSlot(object)
    def handleBatchSaveFinished(self, saved_outputs):
        self.batch_outputs = saved_outputs
        self.refreshBatchOutputsProperty()
        self.progressUpdated.emit(100)
        if self.root:
            self.root.setProperty("processState", "complete")
        QMessageBox.information(None, "Done", f"Saved {len(self.batch_outputs)} files successfully.")
        if self.batch_save_thread:
            self.batch_save_thread.quit()
            self.batch_save_thread.wait()
        self.batch_save_thread = None
        self.batch_save_worker = None

    @pyqtSlot(object,str,str)
    def saveFile(self, df, xml_type, xml_file):
        df = df.replace([float('inf'), float('-inf')], 0).fillna(0)
        dialog = QFileDialog()
        dialog.setWindowTitle("Save Excel File")
        dialog.setNameFilter("Excel Files (*.xlsx)")
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        if self.last_save_dir and os.path.isdir(self.last_save_dir):
            dialog.setDirectory(self.last_save_dir)
        save_path = dialog.selectedFiles()[0] if dialog.exec() else None
        if save_path and not save_path.lower().endswith(".xlsx"): save_path += ".xlsx"
        if not save_path:
            self.thread.quit()
            self.thread.wait()
            self.resetProperties()
            return
        self.rememberSaveDirectory(os.path.dirname(save_path), batch=False)
        if not confirm_overwrite_paths([save_path], "Confirm Overwrite"):
            self.thread.quit()
            self.thread.wait()
            self.resetProperties()
            return

        self.progressUpdated.emit(90)
        self.save_thread = QThread()
        self.save_worker = SaveWorker(df, xml_type, save_path, xml_file)
        self.save_worker.moveToThread(self.save_thread)
        self.save_worker.progress.connect(self.progressUpdated)
        self.save_worker.error.connect(self.handleSaveError)
        self.save_worker.saved.connect(self.handleSaved)
        self.save_thread.started.connect(self.save_worker.save)
        self.save_thread.start()
        self.thread.quit()
        self.thread.wait()

    def handleSaveError(self, msg):
        QMessageBox.critical(None, "Error", msg)
        if self.root:
            self.root.setProperty("processState", "idle")
        self.save_thread.quit()
        self.save_thread.wait()

    @pyqtSlot(str, str)
    def handleSaved(self, save_path, xml_file):
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
        self.save_thread.quit()
        self.save_thread.wait()

    def resetProperties(self):
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
    app.setStyle("Fusion")
    engine=QQmlApplicationEngine()

    backend=Backend(engine)
    engine.rootContext().setContextProperty("backend",backend)

    qml_file=os.path.join(os.path.dirname(__file__),"main.qml")
    engine.load(QUrl.fromLocalFile(qml_file))
    if not engine.rootObjects(): sys.exit(-1)

    backend.root=engine.rootObjects()[0]
    backend.applyRememberedSettingsToUI()
    sys.exit(app.exec())
