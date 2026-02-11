import sys
import os
import pandas as pd
import win32com.client as win32
from PyQt6.QtCore import QObject, pyqtSlot, QUrl, pyqtSignal, QThread, Qt
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox
from PyQt6.QtQml import QQmlApplicationEngine

class Worker(QObject):
    progress = pyqtSignal(int)
    error = pyqtSignal(str)
    dataReady = pyqtSignal(object, str)  # For DataFrame and xml_type

    def __init__(self, xml_file, xml_type):
        super().__init__()
        self.xml_file = xml_file
        self.xml_type = xml_type

    @pyqtSlot()
    def process(self):
        try:
            if self.xml_type not in ["Den", "Globe", "Glacier"]:
                self.error.emit("Only Den, Glacier, and Globe are implemented.")
                return

            xml_file = self.xml_file
            excel = None
            xlsx_file = None
            try:
                excel = win32.DispatchEx('Excel.Application')
                excel.Visible = False
                wb = excel.Workbooks.Open(os.path.abspath(xml_file))
                xlsx_file = os.path.splitext(xml_file)[0] + "_temp.xlsx"
                wb.SaveAs(os.path.abspath(xlsx_file), FileFormat=51)
                wb.Close(False)
            except Exception as e:
                self.error.emit(f"Error converting XML: {e}")
                return
            finally:
                try:
                    excel.Quit()
                except:
                    pass

            self.progress.emit(50)

            # Call the specific processing method based on xml_type
            if self.xml_type == "Den":
                self.process_den(xlsx_file)
            elif self.xml_type == "Globe":
                self.process_globe(xlsx_file)
            elif self.xml_type == "Glacier":
                self.process_glacier(xlsx_file)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred: {e}")

    # -------------------- Den --------------------
    def process_den(self, xlsx_file):
        try:
            df = pd.read_excel(xlsx_file, header=None, engine="openpyxl")
            df = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            header_den_1 = ["0-0:1.0.0", "0-0:96.240.12 [hex]", "1-1:1.5.0 [kW]", "",
                            "1-1:1.8.0 [Wh]", "", "1-1:2.8.0 [Wh]", "1-1:3.8.0 [varh]", "",
                            "1-1:4.8.0 [varh]", "1-1:15.8.1 [Wh]", "1-1:13.5.0", "1-1:128.8.0 [Wh]"]
            header_den_2 = ["Clock", "EDIS status", "Last average demand +A (QI+QIV)", "",
                            "Active energy import +A (QI+QIV)", "", "Active energy export -A (QII+QIII)",
                            "Reactive energy import +R (QI+QII)", "", "Reactive energy export -R (QIII+QIV)",
                            "Active energy A (QI+QII+QIII+QIV) rate 1", "Last average power factor",
                            "Energy |AL1|+|AL2|+|AL3|"]

            # Columns indices
            col_den_ad, col_den_t, col_den_v, col_den_w, col_den_x, col_den_y, col_den_z, col_den_aa, col_den_ab, col_den_ac, col_den_u = 29, 19, 21, 22, 23, 24, 25, 26, 27, 28, 20
            df_filtered = df[df.iloc[:, col_den_ad].notna()].reset_index(drop=True)

            data_den_rows = []
            total_rows = len(df_filtered)
            for i in range(total_rows):
                row = [""] * 13
                row[0] = df_filtered.iloc[i, col_den_t]
                row[1] = df_filtered.iloc[i, col_den_v]
                row[2] = df_filtered.iloc[i, col_den_w]
                row[3] = f"=C{i+3}*280" if i > 0 else ""
                row[4] = df_filtered.iloc[i, col_den_x]
                row[5] = f"=(E{i+3}-E{i+2})*280/1000" if i > 0 else ""
                row[6] = df_filtered.iloc[i, col_den_y]
                row[7] = df_filtered.iloc[i, col_den_z]
                row[8] = f"=(H{i+3}-H{i+2})*280/1000" if i > 0 else ""
                row[9] = df_filtered.iloc[i, col_den_aa]
                row[10] = df_filtered.iloc[i, col_den_ab]
                row[11] = df_filtered.iloc[i, col_den_ac]
                row[12] = df_filtered.iloc[i, col_den_u]
                data_den_rows.append(row)
                progress_value = 50 + int((i / total_rows) * 40)
                self.progress.emit(progress_value)

            final_df = pd.DataFrame([header_den_1, header_den_2] + data_den_rows)

            if xlsx_file and os.path.exists(xlsx_file):
                os.remove(xlsx_file)

            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred in Den processing: {e}")

    # -------------------- Globe --------------------
    def process_globe(self, xlsx_file):
        try:
            df = pd.read_excel(xlsx_file, header=None, engine="openpyxl")
            df = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            header_globe_1 = ["0-0:1.0.0", "0-0:96.240.12 [hex]", "1-1:1.5.0 [kW]", "", "1-1:1.8.0 [Wh]",
                              "", "1-1:1.29.0 [Wh]", "1-1:2.8.0 [Wh]", "1-1:2.29.0 [Wh]", "1-1:3.8.0 [varh]",
                              "", "1-1:3.29.0 [varh]", "1-1:4.8.0 [varh]", "1-1:4.29.0 [varh]", "1-1:13.5.0"]
            header_globe_2 = ["Clock", "EDIS status", "Last average demand +A (QI+QIV)", "",
                              "Active energy import +A (QI+QIV)", "", "Energy delta over capture period 1 +A (QI+QIV)",
                              "Active energy export -A (QII+QIII)", "Energy delta over capture period 1 -A (QII+QIII)",
                              "Reactive energy import +R (QI+QII)", "", "Energy delta over capture period 1 +R (QI+QII)",
                              "Reactive energy export -R (QIII+QIV)", "Energy delta over capture period 1 -R (QIII+QIV)",
                              "Last average power factor"]

            # Indices
            col_globe_af, col_globe_t, col_globe_x, col_globe_y, col_globe_z, col_globe_aa, col_globe_ab, col_globe_ac, col_globe_ad, col_globe_ae, col_globe_u, col_globe_v, col_globe_w = 31, 19, 23, 24, 25, 26, 27, 28, 29, 30, 20, 21, 22
            df_filtered = df[df.iloc[:, col_globe_af].notna()].reset_index(drop=True)

            data_globe_rows = []
            total_rows = len(df_filtered)
            for i in range(total_rows):
                row = [""] * 15
                row[0] = df_filtered.iloc[i, col_globe_t]
                row[1] = df_filtered.iloc[i, col_globe_x]
                row[2] = df_filtered.iloc[i, col_globe_y]
                row[3] = f"=C{i+3}*1400" if i > 0 else ""
                row[4] = df_filtered.iloc[i, col_globe_z]
                row[5] = f"=(E{i+3}-E{i+2})*1400/1000" if i > 0 else ""
                row[6] = df_filtered.iloc[i, col_globe_aa]
                row[7] = df_filtered.iloc[i, col_globe_ab]
                row[8] = df_filtered.iloc[i, col_globe_ac]
                row[9] = df_filtered.iloc[i, col_globe_ad]
                row[10] = f"=(J{i+3}-J{i+2})*1400/1000" if i > 0 else ""
                row[11] = df_filtered.iloc[i, col_globe_ae]
                row[12] = df_filtered.iloc[i, col_globe_u]
                row[13] = df_filtered.iloc[i, col_globe_v]
                row[14] = df_filtered.iloc[i, col_globe_w]
                data_globe_rows.append(row)
                progress_value = 50 + int((i / total_rows) * 40)
                self.progress.emit(progress_value)

            final_df = pd.DataFrame([header_globe_1, header_globe_2] + data_globe_rows)
            if xlsx_file and os.path.exists(xlsx_file):
                os.remove(xlsx_file)

            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred in Globe processing: {e}")

    # -------------------- Glacier --------------------
    def process_glacier(self, xlsx_file):
        try:
            df = pd.read_excel(xlsx_file, header=None, engine="openpyxl")
            df = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            # Example headers, replace with your Glacier logic
            header_glacier_1 = ["Header1", "Header2", "Header3"]
            header_glacier_2 = ["Clock", "Data1", "Data2"]

            data_glacier_rows = []
            for i in range(len(df)):
                row = [df.iloc[i, 0], df.iloc[i, 1] if df.shape[1]>1 else "", df.iloc[i, 2] if df.shape[1]>2 else ""]
                data_glacier_rows.append(row)
                self.progress.emit(50 + int((i / len(df)) * 40))

            final_df = pd.DataFrame([header_glacier_1, header_glacier_2] + data_glacier_rows)

            if xlsx_file and os.path.exists(xlsx_file):
                os.remove(xlsx_file)

            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred in Glacier processing: {e}")


class Backend(QObject):
    progressUpdated = pyqtSignal(int)

    def __init__(self, engine):
        super().__init__()
        self.engine = engine
        self.root = None
        self.selected_file = None
        self.xml_type = ""  # Changed to "" to avoid triggering on init
        self.progress = 0
        self.progressUpdated.connect(self.updateProgressInQML)
        self.thread = None
        self.worker = None

    @pyqtSlot()
    def selectFile(self):
        file_path, _ = QFileDialog.getOpenFileName(None, "Select XML File", "", "XML Files (*.xml)")
        if not file_path:
            QMessageBox.information(None, "Info", "No file selected")
            return
        self.selected_file = file_path
        if self.root:
            self.root.setProperty("selectedFile", file_path)
            # Calculate and set actual file size
            try:
                size_bytes = os.path.getsize(file_path)
                # Format size: KB if < 1 MB, MB otherwise
                if size_bytes < 1024 * 1024:
                    size_str = f"{size_bytes / 1024:.2f} KB"
                else:
                    size_str = f"{size_bytes / (1024 * 1024):.2f} MB"
                self.root.setProperty("fileSize", size_str)
            except OSError:
                self.root.setProperty("fileSize", "Unknown")
            self.root.setProperty("processState", "selecting")

    @pyqtSlot(str)
    def setSelectionType(self, type_str):
        self.xml_type = type_str
        if type_str == "Glacier":
            QMessageBox.information(None, "Glacier", "Glacier XML processing not implemented yet!")
        # Removed message for Globe since it's now implemented
        if self.root:
            self.root.setProperty("selectionType", type_str)

    @pyqtSlot()
    def confirmAndConvert(self):
        if not self.xml_type or self.xml_type == "":
            return
        if self.root:
            self.root.setProperty("processState", "converting")
        self.progress = 0
        self.progressUpdated.emit(self.progress)
        # Start processing in a separate thread
        self.thread = QThread()
        self.worker = Worker(self.selected_file, self.xml_type)
        self.worker.moveToThread(self.thread)
        self.worker.progress.connect(self.progressUpdated)
        self.worker.error.connect(self.handleError)
        self.worker.dataReady.connect(self.saveFile)
        self.thread.started.connect(self.worker.process)
        self.thread.start()

    @pyqtSlot()
    def selectDifferentFile(self):
        self.resetProperties()

    @pyqtSlot(str, result=str)  # New slot for file size calculation (used by QML on drop)
    def getFileSize(self, file_path):
        try:
            size_bytes = os.path.getsize(file_path)
            if size_bytes < 1024 * 1024:
                return f"{size_bytes / 1024:.2f} KB"
            else:
                return f"{size_bytes / (1024 * 1024):.2f} MB"
        except OSError:
            return "Unknown"        

    @pyqtSlot()
    def convertAnotherFile(self):
        self.resetProperties()

    @pyqtSlot(str)
    def setSelectedFile(self, file_path):
        self.selected_file = file_path

    def updateProgressInQML(self, value):
        if self.root:
            self.root.setProperty("progress", value)

    def handleError(self, msg):
        QMessageBox.critical(None, "Error", msg)
        if self.root:
            self.root.setProperty("processState", "idle")
        self.thread.quit()
        self.thread.wait()

    @pyqtSlot(object, str)
    def saveFile(self, df, xml_type):
        dialog = QFileDialog()
        dialog.setWindowTitle("Save Excel File")
        dialog.setNameFilter("Excel Files (*.xlsx)")
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        if dialog.exec():
            save_path = dialog.selectedFiles()[0]
        else:
            save_path = None

        if not save_path:
            if self.root:
                self.root.setProperty("processState", "idle")
            self.thread.quit()
            self.thread.wait()
            return

        self.progressUpdated.emit(90)

        try:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False)
                workbook = writer.book
                ws = writer.sheets['Sheet1']

                # ===== Formats =====
                general_fmt = workbook.add_format({'num_format': 'General', 'border': 1, 'align': 'right'})
                text_fmt = workbook.add_format({'num_format': '@', 'border': 1, 'align': 'right'})
                num_fmt = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'right'})
                header_fmt = workbook.add_format({
                    'num_format': '@', 'bg_color': '#99CC00', 'font_color': 'white',
                    'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1
                })

                if xml_type == "Globe":
                    colored_fmt = workbook.add_format({'bg_color': '#B4C6E7', 'border': 1, 'align': 'right'})
                    colored_header_fmt = workbook.add_format({
                        'num_format': '@', 'bg_color': '#B4C6E7', 'font_color': 'white',
                        'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1
                    })
                    colored_num_fmt = workbook.add_format({
                        'num_format': '0.00', 'bg_color': '#B4C6E7', 'border': 1, 'align': 'right'
                    })

                # ===== Write headers =====
                for r in range(2):  # Rows 0 and 1 → headers
                    for c in range(df.shape[1]):
                        val = df.iloc[r, c]
                        if xml_type == "Globe" and c in [3, 5, 10]:
                            ws.write(r, c, val, colored_header_fmt)
                        else:
                            ws.write(r, c, val, header_fmt)

                # ===== Write data starting from row 3 =====
                for r in range(2, len(df)):
                    for c in range(df.shape[1]):
                        val = df.iloc[r, c]
                        if xml_type == "Den":
                            # Force Column B (index 1) as text
                            if c == 1:
                                ws.write(r, c, val, text_fmt)
                            elif isinstance(val, str) and val.startswith("="):
                                ws.write_formula(r, c, val, num_fmt)  # formulas right-aligned
                            elif isinstance(val, (int, float)):
                                ws.write(r, c, val, num_fmt)
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
                        else:  # Glacier default
                            ws.write(r, c, val, general_fmt)

                # ===== Column widths =====
                if xml_type == "Den":
                    widths = [17.73, 17.27] + [16.27]*7 + [32.27, 36.36, 23.36, 24.76]
                    for i, w in enumerate(widths):
                        ws.set_column(i, i, w)
                    ws.set_column(6, 6, 16.27, num_fmt, options={'hidden': True}) #Hides G

                elif xml_type == "Globe":
                    ws.set_column(0, 0, 17.73)
                    ws.set_column(1, 1, 17.27)
                    ws.set_column(2, 10, 14.91)
                    ws.set_column(11, 11, 41.91)
                    ws.set_column(12, 12, 33.27)
                    ws.set_column(13, 13, 43.36)
                    ws.set_column(14, 14, 23.36)
                    ws.set_column(6, 8, None, None, {'hidden': True})  # Hide G-I

                elif xml_type == "Glacier":
                    for c in range(df.shape[1]):
                        ws.set_column(c, c, 17)

        except Exception as e:
            QMessageBox.critical(None, "Error", f"Failed to save Excel: {e}")
            if self.root:
                self.root.setProperty("processState", "idle")
            self.thread.quit()
            self.thread.wait()
            return

        self.progressUpdated.emit(100)
        if self.root:
            self.root.setProperty("processState", "complete")
        QMessageBox.information(None, "Done", f"Processed Excel saved:\n{save_path}")
        self.thread.quit()
        self.thread.wait()


    def resetProperties(self):
        if self.root:
            self.root.setProperty("processState", "idle")
            self.root.setProperty("selectedFile", "")
            self.root.setProperty("selectionType", "")
            self.root.setProperty("fileSize", "")
            self.root.setProperty("progress", 0)
        self.selected_file = None
        self.xml_type = ""
        self.progress = 0

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Set to Fusion style to allow QML customizations
    engine = QQmlApplicationEngine()

    backend = Backend(engine)
    engine.rootContext().setContextProperty("backend", backend)

    qml_file = os.path.join(os.path.dirname(__file__), "main.qml")
    engine.load(QUrl.fromLocalFile(qml_file))

    if not engine.rootObjects():
        sys.exit(-1)

    backend.root = engine.rootObjects()[0]

    sys.exit(app.exec())
