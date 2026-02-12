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
            xml_file = self.xml_file

            if self.xml_type not in ["Den", "Globe", "Glacier"]:
                self.error.emit("Only Den, Glacier, and Globe are implemented.")
                return

            # --------- Use pandas.read_xml instead of Excel conversion ---------
            try:
                df_xml = pd.read_xml(
                    xml_file,
                    xpath=".//ns:Items",
                    namespaces={"ns": "http://tempuri.org/ArrayFieldDataSet.xsd"}
                )
            except Exception as e:
                self.error.emit(f"Error reading XML directly: {e}")
                return

            self.progress.emit(50)

            # Call existing processing methods, but pass df directly
            if self.xml_type == "Den":
                self.process_den_from_df(df_xml)
            elif self.xml_type == "Globe":
                self.process_globe_from_df(df_xml)
            elif self.xml_type == "Glacier":
                self.process_glacier_from_df(df_xml)

        except Exception as e:
            self.error.emit(f"An unexpected error occurred: {e}")


    # -------------------- Den --------------------
    def process_den_from_df(self, df):
        try:
            # Filter out irrelevant rows
            df_filtered = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            if df_filtered.empty:
                self.error.emit("No valid rows found in Den XML.")
                return

            # Take only columns C2-C10 (original columns 1-9)
            df_data = df_filtered.iloc[:, 1:12]  #df_data columns 0-8 = C2-C10


            # Define column mapping: df_data index → Excel column index
            column_mapping = {
                0: 0,  # C1 → A
                1: 1,  # C2 → B
                2: 2,  # C3 → C
                3: 4,  # C4 → E
                4: 6,  # C5 → G
                5: 7,  # C6 → H
                6: 9,  # C7 → J
                7: 10, # C8 → K
                8: 11,  # C9 → L   
                9: 12   # C10 → M                      
            }

            # Create an empty final DataFrame with enough columns
            max_col = max(column_mapping.values()) + 1
            final_rows = []

            # Add headers
            header_den_row_1 = [""] * max_col
            header_den_row_2 = [""] * max_col

            header_den_map_1 = {
                0: "0-0:1.0.0", 1: "0-0:96.240.12 [hex]", 2: "1-1:1.5.0 [kW]",
                4: "1-1:1.8.0 [Wh]", 6: "1-1:2.8.0 [Wh]", 7: "1-1:3.8.0 [varh]",
                9: "1-1:4.8.0 [varh]", 10: "1-1:15.8.1 [Wh]", 11: "1-1:13.5.0", 12: "1-1:128.8.0 [Wh]"
            }

            header_den_map_2 = {
                0: "Clock", 1: "EDIS status", 2: "Last average demand +A (QI+QIV)",
                4: "Active energy import +A (QI+QIV)", 6: "Active energy export -A (QII+QIII)",
                7: "Reactive energy import +R (QI+QII)", 9: "Reactive energy export -R (QIII+QIV)",
                10: "Active energy A (QI+QII+QIII+QIV) rate 1", 11: "Last average power factor",
                12: "Energy |AL1|+|AL2|+|AL3|"
            }

            for idx, val in header_den_map_1.items():
                header_den_row_1[idx] = val
            for idx, val in header_den_map_2.items():
                header_den_row_2[idx] = val

            final_rows = [header_den_row_1,header_den_row_2]

            for i in range(len(df_data)):
                row = [""] * max_col  # max_col = highest Excel column index you will use + 1

                # Fill data according to your mapping
                for idx, col in column_mapping.items():
                    row[col] = df_data.iloc[i, idx]

                # Add formulas
                if i > 0:
                    row[3] = f"=C{i+3}*280"          # Column D
                    row[5] = f"=(E{i+3}-E{i+2})*280/1000"  # Column F
                    row[8] = f"=(H{i+3}-H{i+2})*280/1000"  # Column I

                final_rows.append(row)

            final_df = pd.DataFrame(final_rows)


            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)

        except Exception as e:
            self.error.emit(f"An unexpected error occurred in Den processing: {e}")

    # -------------------- Globe --------------------
    def process_globe_from_df(self, df):
        try:
            df_filtered = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            if df_filtered.empty:
                self.error.emit("No valid rows found in Globe XML.")
                return

            df_data = df_filtered.iloc[:, 1:14]

            column_mapping = {
                0: 0,   
                1: 1,   
                2: 2, 
                3: 4, 
                4: 6,  
                5: 7,  
                6: 8,  
                7: 9, 
                8: 11,  
                9: 12, 
                10: 13,  
                11: 14, 
            }
            max_col = max(column_mapping.values()) + 1

            header_globe_row_1 = [""] * max_col
            header_globe_row_2 = [""] * max_col

            header_map_1 = {
                0: "0-0:1.0.0", 1: "0-0:96.240.12 [hex]", 2: "1-1:1.5.0 [kW]",
                4: "1-1:1.8.0 [Wh]", 6: "1-1:1.29.0 [Wh]", 7: "1-1:2.8.0 [Wh]",
                8: "1-1:2.29.0 [Wh]", 9: "1-1:3.8.0 [varh]", 11: "1-1:3.29.0 [varh]",
                12: "1-1:4.8.0 [varh]",13: "1-1:4.29.0 [varh]", 14: "1-1:13.5.0" }

            header_map_2 = {
                0: "Clock", 1: "EDIS status", 2: "Last average demand +A (QI+QIV)", 
                4: "Active energy import +A (QI+QIV)", 6: "Energy delta over capture period 1 +A (QI+QIV)",
                7: "Active energy export -A (QII+QIII)", 8: "Energy delta over capture period 1 -A (QII+QIII)",
                9: "Reactive energy import +R (QI+QII)", 11: "Energy delta over capture period 1 +R (QI+QII)",
                12: "Reactive energy export -R (QIII+QIV)", 13: "Energy delta over capture period 1 -R (QIII+QIV)",
                14: "Last average power factor"
            }
            for idx, val in header_map_1.items():
                header_globe_row_1[idx] = val
            for idx, val in header_map_2.items():
                header_globe_row_2[idx] = val

            final_rows = [header_globe_row_1, header_globe_row_2]

            total_rows = len(df_data)
            for i in range(total_rows):
                row = [""] * max_col
                for idx, col in column_mapping.items():
                    row[col] = df_data.iloc[i, idx]

                if i > 0:
                    row[3] = f"=C{i+3}*1400"      
                    row[5] = f"=(E{i+3}-E{i+2})*1400/1000"  
                    row[10] = f"=(J{i+3}-J{i+2})*1400/1000" 

                final_rows.append(row)
                progress_value = 50 + int((i / total_rows) * 40)
                self.progress.emit(progress_value)

            final_df = pd.DataFrame(final_rows)

            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)

        except Exception as e:
            self.error.emit(f"An unexpected error occurred in Globe processing: {e}")


# -------------------- Glacier --------------------
    def process_glacier_from_df(self, df):
        try:
            df_filtered = df[~df.iloc[:, 0].astype(str).str.contains("/ArrayFieldDataSet", na=False)].reset_index(drop=True)

            if df_filtered.empty:
                self.error.emit("No valid rows found in Glacier XML.")
                return
            df_data = df_filtered.iloc[:, 1:12]

            column_mapping = {
                0: 0,  
                1: 1,  
                2: 2,  
                3: 4, 
                4: 6,  
                5: 7,   
                6: 9,  
                7: 10, 
                8: 11, 
                9: 12 
            }
            max_col = max(column_mapping.values()) + 1

            header_row_1 = [""] * max_col
            header_row_2 = [""] * max_col

            header_map_1 = {
                0: "0-0:1.0.0", 1: "0-0:96.240.12 [hex]", 2: "1-1:1.5.0 [kW]",
                4: "1-1:1.8.0 [Wh]", 6: "1-1:2.8.0 [Wh]", 7: "1-1:3.8.0 [varh]",
                9: "1-1:4.8.0 [varh]", 10: "1-1:15.8.1 [Wh]", 11: "1-1:13.5.0", 12: "1-1:128.8.0 [Wh]"
            }
            header_map_2 = {
                0: "Clock", 1: "EDIS status", 2: "Last average demand +A (QI+QIV)",
                4: "Active energy import +A (QI+QIV)", 6: "Active energy export -A (QII+QIII)",
                7: "Reactive energy import +R (QI+QII)", 9: "Reactive energy export -R (QIII+QIV)",
                10: "Active energy A (QI+QII+QIII+QIV) rate 1", 11: "Last average power factor",
                12: "Energy |AL1|+|AL2|+|AL3|"
            }
            for idx, val in header_map_1.items():
                header_row_1[idx] = val
            for idx, val in header_map_2.items():
                header_row_2[idx] = val

            final_rows = [header_row_1, header_row_2]

            for i in range(len(df_data)):
                row = [""] * max_col
                for idx, col in column_mapping.items():
                    row[col] = df_data.iloc[i, idx]

                if i > 0:
                    row[3] = f"=C{i+3}*280"              
                    row[5] = f"=(E{i+3}-E{i+2})*280/1000"  
                    row[8] = f"=(H{i+3}-H{i+2})*280/1000" 

                final_rows.append(row)

            final_df = pd.DataFrame(final_rows)

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
        self.xml_type = "" 
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

            try:
                size_bytes = os.path.getsize(file_path)

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

    @pyqtSlot(str, result=str)
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
        df = df.replace([float('inf'), float('-inf')], 0).fillna(0)

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
                xml_file_name = os.path.splitext(os.path.basename(self.selected_file))[0]
                parts = xml_file_name.split('_')
                if '_' in xml_file_name:
                    sheet_name = '_'.join(xml_file_name.split('_')[:-1])
                else:
                    sheet_name = xml_file_name
                sheet_name = sheet_name[:31]
                df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
                workbook = writer.book
                ws = writer.sheets[sheet_name]

                # ===== Formats =====
                general_fmt = workbook.add_format({'num_format': 'General', 'border': 1, 'align': 'right'})
                text_fmt = workbook.add_format({'num_format': '@', 'border': 1, 'align': 'right'})
                num_fmt = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'right'})
                header_fmt = workbook.add_format({
                    'num_format': '@', 'bg_color': '#99CC00', 'font_color': 'white',
                    'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1
                })

                # Define colored formats for types that use them
                if xml_type in ["Globe", "Glacier"]:
                    colored_num_fmt = workbook.add_format({
                        'num_format': '0.00', 'bg_color': '#B4C6E7', 'border': 1, 'align': 'right'
                    })
                    colored_header_fmt = workbook.add_format({
                        'num_format': '@', 'bg_color': '#B4C6E7', 'font_color': 'white',
                        'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1
                    })

                if xml_type == "Globe":
                    colored_fmt = workbook.add_format({'bg_color': '#B4C6E7', 'border': 1, 'align': 'right'})

                # ===== Write headers =====
                for r in range(2):  # Rows 0 and 1 → headers
                    for c in range(df.shape[1]):
                        val = df.iloc[r, c]
                        if xml_type == "Globe" and c in [3, 5, 10]:
                            ws.write(r, c, val, colored_header_fmt)
                        elif xml_type in ["Glacier"] and c in [3, 5, 8]:
                            ws.write(r, c, val, colored_header_fmt)
                        else:
                            ws.write(r, c, val, header_fmt)

                # ===== Write data starting from row 2 =====
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
                            if c in [3, 5, 8]:
                                ws.write(r, c, val, colored_num_fmt)
                            elif 3 <= c <= 12 and c not in [3, 5, 8]:  # Adjusted for Glacier's 13 columns
                                ws.write(r, c, val, num_fmt)
                            elif c == 0:
                                ws.write(r, c, val, general_fmt)
                            elif c == 1:
                                ws.write(r, c, val, text_fmt)
                            else:
                                ws.write(r, c, val, num_fmt)

                # ===== Column widths =====
                if xml_type == "Den":
                    widths = [17.73, 17.27] + [16.27]*7 + [32.27, 36.36, 23.36, 24.76]
                    for i, w in enumerate(widths):
                        ws.set_column(i, i, w)
                    ws.set_column(6, 6, 16.27, None,{'hidden': True})

                elif xml_type == "Globe":
                    ws.set_column(0, 0, 17.73)
                    ws.set_column(1, 1, 17.27)
                    ws.set_column(2, 10, 14.91)
                    ws.set_column(11, 11, 41.91)
                    ws.set_column(12, 12, 33.27)
                    ws.set_column(13, 13, 43.36)
                    ws.set_column(14, 14, 23.36)
                    ws.set_column(6, 8, 14.91, None, {'hidden': True})

                elif xml_type == "Glacier":
                    ws.set_column(0, 0, 17.73)
                    ws.set_column(1, 1, 17.27)
                    ws.set_column(2, 8, 16.91)
                    ws.set_column(9, 9, 18.73)
                    ws.set_column(10, 10, 17.55)
                    ws.set_column(11, 11, 23.36)
                    ws.set_column(12, 12, 23.36) 
                    ws.set_column(6, 6, 16.91, None, {'hidden': True}) 

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
