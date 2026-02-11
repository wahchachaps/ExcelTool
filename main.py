import sys
import os
import pandas as pd
import win32com.client as win32
import time
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
            if self.xml_type not in ["Den", "Globe"]:
                self.error.emit("Only Den and Globe are implemented.")
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
            else:
                self.error.emit("Unsupported XML type.")
        except Exception as e:
            self.error.emit(f"An unexpected error occurred: {e}")

    def process_den(self, xlsx_file):
        try:
            try:
                df = pd.read_excel(xlsx_file, header=None, engine="openpyxl")
            except Exception as e:
                self.error.emit(f"Failed to read Excel file: {e}")

                if xlsx_file and os.path.exists(xlsx_file):
                    os.remove(xlsx_file)
                return

            df = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            raw_den_headers = ["0-0:1.0.0", "0-0:96.240.12 [hex]", "1-1:1.5.0 [kW]", "",
                               "1-1:1.8.0 [Wh]", "", "1-1:2.8.0 [Wh]", "1-1:3.8.0 [varh]", "",
                               "1-1:4.8.0 [varh]", "1-1:15.8.1 [Wh]", "1-1:13.5.0", "1-1:128.8.0 [Wh]"]
            friendly_den_headers = ["Clock", "EDIS status", "Last average demand +A (QI+QIV)", "Demand",
                                    "Active energy import +A (QI+QIV)", "kWh", "Active energy export -A (QII+QIII)",
                                    "Reactive energy import +R (QI+QII)", "kVarh", "Reactive energy export -R (QIII+QIV)",
                                    "Active energy A (QI+QII+QIII+QIV) rate 1", "Last average power factor",
                                    "Energy |AL1|+|AL2|+|AL3|"]

            col_den_ad, col_den_t, col_den_v, col_den_w, col_den_x, col_den_y, col_den_z, col_den_aa, col_den_ab, col_den_ac, col_den_u = 29, 19, 21, 22, 23, 24, 25, 26, 27, 28, 20
            df_filtered = df[df.iloc[:, col_den_ad].notna()].reset_index(drop=True)

            clock_data = df_filtered.iloc[:, col_den_t]
            edis_data = df_filtered.iloc[:, col_den_v]
            col_c_data = df_filtered.iloc[:, col_den_w]
            col_e_data = df_filtered.iloc[:, col_den_x]
            col_g_data = df_filtered.iloc[:, col_den_y]
            col_i_data = df_filtered.iloc[:, col_den_z]
            col_j_data = df_filtered.iloc[:, col_den_aa]
            col_k_data = df_filtered.iloc[:, col_den_ab]
            col_l_data = df_filtered.iloc[:, col_den_ac]
            col_m_data = df_filtered.iloc[:, col_den_u]

            data_den_rows = []
            total_rows = len(df_filtered)
            for i in range(total_rows):
                row = [""] * 13
                row[0] = clock_data[i]
                row[1] = edis_data[i]
                row[2] = col_c_data[i]
                row[3] = f"=C{i+3}*280" if i > 0 else ""
                row[4] = col_e_data[i]
                row[5] = f"=(E{i+3}-E{i+2})*280/1000" if i > 0 else ""
                row[6] = col_g_data[i]
                row[7] = col_i_data[i]
                row[8] = f"=(H{i+3}-H{i+2})*280/1000" if i > 0 else ""
                row[9] = col_j_data[i]
                row[10] = col_k_data[i]
                row[11] = col_l_data[i]
                row[12] = col_m_data[i]
                data_den_rows.append(row)
                # Smoother progress updates (from 50% to 90%)
                progress_value = 50 + int((i / total_rows) * 40)
                self.progress.emit(progress_value)

            final_df = pd.DataFrame([raw_den_headers, friendly_den_headers] + data_den_rows)


            if xlsx_file and os.path.exists(xlsx_file):
                os.remove(xlsx_file)

            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred in Den processing: {e}")

    def process_globe(self, xlsx_file):
        try:
            try:
                df = pd.read_excel(xlsx_file, header=None, engine="openpyxl")
            except Exception as e:
                self.error.emit(f"Failed to read Excel file: {e}")

                if xlsx_file and os.path.exists(xlsx_file):
                    os.remove(xlsx_file)
                return

            df = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            raw_globe_headers = ["", "", "", "", "", "", "1-1:1.29.0 [Wh]", "1-1:2.8.0 [Wh]", "1-1:2.29.0 [Wh]", "1-1:3.8.0 [varh]",
                                 "", "1-1:3.29.0 [varh]", "1-1:4.8.0 [varh]", "1-1:4.29.0 [varh]", "1-1:13.5.0"]
            friendly_globe_headers = ["Clock", "EDIS status", "Last average demand +A (QI+QIV)", "",
                                      "Active energy import +A (QI+QIV)", "", "Energy delta over capture period 1 +A (QI+QIV)",
                                      "Active energy export -A (QII+QIII)", "Energy delta over capture period 1 -A (QII+QIII)",
                                      "Reactive energy import +R (QI+QII)", "", "Energy delta over capture period 1 +R (QI+QII)",
                                      "Reactive energy export -R (QIII+QIV)", "Energy delta over capture period 1 -R (QIII+QIV)",
                                      "Last average power factor"]

            col_globe_af, col_globe_t, col_globe_x, col_globe_y, col_globe_z, col_globe_aa, col_globe_ab, col_globe_ac, col_globe_ad, col_globe_ae, col_globe_u, col_globe_v, col_globe_w = 31, 19, 23, 24, 25, 26, 27, 28, 29, 30, 20, 21, 22
            df_filtered = df[df.iloc[:, col_globe_af].notna()].reset_index(drop=True)

            col_a_data = df_filtered.iloc[:, col_globe_t]
            col_b_data = df_filtered.iloc[:, col_globe_x]
            col_c_data = df_filtered.iloc[:, col_globe_y]
            col_e_data = df_filtered.iloc[:, col_globe_z]
            col_g_data = df_filtered.iloc[:, col_globe_aa]
            col_h_data = df_filtered.iloc[:, col_globe_ab]
            col_i_data = df_filtered.iloc[:, col_globe_ac]
            col_j_data = df_filtered.iloc[:, col_globe_ad]
            col_l_data = df_filtered.iloc[:, col_globe_ae]
            col_m_data = df_filtered.iloc[:, col_globe_u]
            col_n_data = df_filtered.iloc[:, col_globe_v]
            col_o_data = df_filtered.iloc[:, col_globe_w]

            data_globe_rows = []
            total_rows = len(df_filtered)
            for i in range(total_rows):
                row = [""] * 15  # Fixed to match the number of headers and assignments
                row[0] = col_a_data[i]
                row[1] = col_b_data[i]
                row[2] = col_c_data[i]
                row[3] = f"=C{i+3}*280" if i > 0 else ""
                row[4] = col_e_data[i]
                row[5] = f"=(E{i+3}-E{i+2})*280/1000" if i > 0 else ""
                row[6] = col_g_data[i]
                row[7] = col_h_data[i]
                row[8] = col_i_data[i]
                row[9] = col_j_data[i]
                row[10] = f"=(J{i+3}-J{i+2})*280/1000" if i > 0 else ""
                row[11] = col_l_data[i]
                row[12] = col_m_data[i]
                row[13] = col_n_data[i]
                row[14] = col_o_data[i]
                data_globe_rows.append(row)
                # Smoother progress updates (from 50% to 90%)
                progress_value = 50 + int((i / total_rows) * 40)
                self.progress.emit(progress_value)

            final_df = pd.DataFrame([raw_globe_headers, friendly_globe_headers] + data_globe_rows)

            if xlsx_file and os.path.exists(xlsx_file):
                os.remove(xlsx_file)

            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred in Globe processing: {e}")

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

        self.progressUpdated.emit(90)  # Already at 90, but ensure
        try:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False)
                workbook = writer.book
                ws = writer.sheets['Sheet1']
                general_fmt = workbook.add_format({'num_format': 'General', 'border': 1, 'align': 'right'})
                text_fmt = workbook.add_format({'num_format': '@', 'border': 1, 'align': 'right'})
                num_fmt = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'right'})
                header_fmt = workbook.add_format({
                    'bg_color': '#99CC00', 'font_color': 'white',
                    'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1
                })
                if xml_type == "Globe":
                    colored_fmt = workbook.add_format({'bg_color': '#B4C6E7', 'border': 1, 'align': 'right'})
                    colored_header_fmt = workbook.add_format({
                        'bg_color': '#B4C6E7', 'font_color': 'white',
                        'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1
                    })

                # Write headers with appropriate formats
                for r in range(2):
                    for c in range(df.shape[1]):
                        val = df.iloc[r, c]
                        if xml_type == "Globe" and c in [3, 5, 10]:  # Columns D, F, K
                            ws.write(r, c, val, colored_header_fmt)
                        else:
                            ws.write(r, c, val, header_fmt)

                for r in range(2, len(df)):
                    for c in range(df.shape[1]):
                        val = df.iloc[r, c]
                        if xml_type == "Globe" and c in [3, 5, 10]:  # Columns D, F, K
                            ws.write(r, c, val, colored_fmt)
                        elif c in [0, 1]:
                            ws.write(r, c, val, general_fmt)
                        elif c in [2, 3, 5, 8]:
                            ws.write(r, c, val, text_fmt)
                        else:
                            ws.write(r, c, val, num_fmt)

                # Adjust column widths based on the maximum length in the entire column (headers and data)
                for c in range(df.shape[1]):
                    max_len = max(len(str(df.iloc[r, c])) for r in range(len(df)))
                    ws.set_column(c, c, max_len + 5)
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