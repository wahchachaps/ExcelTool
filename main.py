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
    dataReady = pyqtSignal(object, str)

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
            try:
                df_xml = pd.read_xml(
                    self.xml_file,
                    xpath=".//ns:Items",
                    namespaces={"ns": "http://tempuri.org/ArrayFieldDataSet.xsd"}
                )
            except Exception as e:
                self.error.emit(f"Error reading XML directly: {e}")
                return
            self.progress.emit(50)
            if self.xml_type == "Den": self.process_den_from_df(df_xml)
            elif self.xml_type == "Globe": self.process_globe_from_df(df_xml)
            elif self.xml_type == "Glacier": self.process_glacier_from_df(df_xml)
        except Exception as e:
            self.error.emit(f"An unexpected error occurred: {e}")

    def process_den_from_df(self, df):
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
            for idx,val in {0:"Clock",1:"EDIS status",2:"Last average demand +A (QI+QIV)",4:"Active energy import +A (QI+QIV)",6:"Active energy export -A (QII+QIII)",7:"Reactive energy import +R (QI+QII)",9:"Reactive energy export -R (QIII+QIV)",10:"Active energy A (QI+QII+QIII+QIV) rate 1",11:"Last average power factor",12:"Energy |AL1|+|AL2|+|AL3|"}.items(): header_den_row_2[idx]=val
            final_rows = [header_den_row_1, header_den_row_2]
            for i in range(len(df_data)):
                row = [""]*max_col
                for idx,col in column_mapping.items(): row[col]=df_data.iloc[i,idx]
                if i>0:
                    row[3]=f"=C{i+3}*280"; row[5]=f"=(E{i+3}-E{i+2})*280/1000"; row[8]=f"=(H{i+3}-H{i+2})*280/1000"
                final_rows.append(row)
            final_df = pd.DataFrame(final_rows)
            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"Den processing error: {e}")

    def process_globe_from_df(self, df):
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
            for idx,val in {0:"Clock",1:"EDIS status",2:"Last average demand +A (QI+QIV)",4:"Active energy import +A (QI+QIV)",6:"Energy delta over capture period 1 +A (QI+QIV)",7:"Active energy export -A (QII+QIII)",8:"Energy delta over capture period 1 -A (QII+QIII)",9:"Reactive energy import +R (QI+QII)",11:"Energy delta over capture period 1 +R (QI+QII)",12:"Reactive energy export -R (QIII+QIV)",13:"Energy delta over capture period 1 -R (QIII+QIV)",14:"Last average power factor"}.items(): header_globe_row_2[idx]=val
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
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"Globe processing error: {e}")

    def process_glacier_from_df(self, df):
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
            for idx,val in {0:"Clock",1:"EDIS status",2:"Last average demand +A (QI+QIV)",4:"Active energy import +A (QI+QIV)",6:"Active energy export -A (QII+QIII)",7:"Reactive energy import +R (QI+QII)",9:"Reactive energy export -R (QIII+QIV)",10:"Active energy A (QI+QII+QIII+QIV) rate 1",11:"Last average power factor",12:"Energy |AL1|+|AL2|+|AL3|"}.items(): header_row_2[idx]=val
            final_rows = [header_row_1, header_row_2]
            for i in range(len(df_data)):
                row = [""]*max_col
                for idx,col in column_mapping.items(): row[col]=df_data.iloc[i,idx]
                if i>0: row[3]=f"=C{i+3}*280"; row[5]=f"=(E{i+3}-E{i+2})*280/1000"; row[8]=f"=(H{i+3}-H{i+2})*280/1000"
                final_rows.append(row)
            final_df = pd.DataFrame(final_rows)
            self.progress.emit(90)
            self.dataReady.emit(final_df, self.xml_type)
        except Exception as e:
            self.error.emit(f"Glacier processing error: {e}")

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
        file_path,_ = QFileDialog.getOpenFileName(None,"Select XML File","","XML Files (*.xml)")
        if not file_path:
            QMessageBox.information(None,"Info","No file selected")
            return
        self.selected_file = file_path
        if self.root:
            self.root.setProperty("selectedFile",file_path)
            try:
                size_bytes=os.path.getsize(file_path)
                size_str=f"{size_bytes/1024:.2f} KB" if size_bytes<1024*1024 else f"{size_bytes/(1024*1024):.2f} MB"
                self.root.setProperty("fileSize",size_str)
            except OSError:
                self.root.setProperty("fileSize","Unknown")
            self.root.setProperty("processState","selecting")

    @pyqtSlot(str)
    def setSelectionType(self,type_str):
        self.xml_type = type_str
        if self.root:
            self.root.setProperty("selectionType",type_str)

    @pyqtSlot()
    def confirmAndConvert(self):
        if not self.xml_type:
            return
        if self.root:
            self.root.setProperty("processState","converting")
        self.progress=0
        self.progressUpdated.emit(self.progress)
        self.thread=QThread()
        self.worker=Worker(self.selected_file,self.xml_type)
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

    @pyqtSlot()
    def convertAnotherFile(self):
        self.resetProperties()

    @pyqtSlot(str)
    def setSelectedFile(self,file_path):
        self.selected_file=file_path

    def updateProgressInQML(self,value):
        if self.root:
            self.root.setProperty("progress",value)

    def handleError(self,msg):
        QMessageBox.critical(None,"Error",msg)
        if self.root:
            self.root.setProperty("processState","idle")
        self.thread.quit()
        self.thread.wait()

    @pyqtSlot(object,str)
    def saveFile(self,df,xml_type):
        df=df.replace([float('inf'),float('-inf')],0).fillna(0)
        dialog=QFileDialog()
        dialog.setWindowTitle("Save Excel File")
        dialog.setNameFilter("Excel Files (*.xlsx)")
        dialog.setWindowFlags(dialog.windowFlags()|Qt.WindowType.WindowStaysOnTopHint)
        save_path=dialog.selectedFiles()[0] if dialog.exec() else None
        if save_path and not save_path.lower().endswith(".xlsx"): save_path+=".xlsx"
        if not save_path:
            if self.root:
                self.root.setProperty("processState","idle")
            self.thread.quit()
            self.thread.wait()
            return

        self.progressUpdated.emit(90)
        try:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                xml_file_name=os.path.splitext(os.path.basename(self.selected_file))[0]
                sheet_name=('_'.join(xml_file_name.split('_')[:-1]) if '_' in xml_file_name else xml_file_name)[:31]
                df.to_excel(writer,index=False,header=False,sheet_name=sheet_name)
                workbook=writer.book
                ws=writer.sheets[sheet_name]

                general_fmt=workbook.add_format({'num_format':'General','border':1,'align':'right'})
                text_fmt=workbook.add_format({'num_format':'@','border':1,'align':'right'})
                num_fmt=workbook.add_format({'num_format':'0.00','border':1,'align':'right'})
                header_fmt=workbook.add_format({'num_format':'@','bg_color':'#99CC00','font_color':'white','align':'center','valign':'vcenter','left':1,'right':1})
                colored_header_fmt=workbook.add_format({'num_format':'@','bg_color':'#B4C6E7','font_color':'white','align':'center','valign':'vcenter','left':1,'right':1})
                colored_num_fmt=workbook.add_format({'num_format':'0.00','bg_color':'#B4C6E7','border':1,'align':'right'})

                for r in range(2):
                    for c in range(df.shape[1]):
                        val=df.iloc[r,c]
                        if xml_type=="Globe" and c in [3,5,10]: ws.write(r,c,val,colored_header_fmt)
                        elif xml_type in ["Glacier","Den"] and c in [3,5,8]: ws.write(r,c,val,colored_header_fmt)
                        else: ws.write(r,c,val,header_fmt)

                for r in range(2,len(df)):
                    for c in range(df.shape[1]):
                        val=df.iloc[r,c]
                        if xml_type=="Den":
                            if c==1: ws.write(r,c,val,text_fmt)
                            elif c in [3,5,8]:
                                ws.write_formula(r,c,val,colored_num_fmt) if isinstance(val,str) and val.startswith("=") else ws.write(r,c,val,colored_num_fmt)
                            elif isinstance(val,(int,float)): ws.write(r,c,val,num_fmt)
                            elif isinstance(val,str) and val.startswith("="): ws.write_formula(r,c,val,num_fmt)
                            else: ws.write(r,c,val,text_fmt)
                        elif xml_type=="Globe":
                            if c in [3,5,10]: ws.write(r,c,val,colored_num_fmt)
                            elif 3<=c<=14 and c not in [3,5,10]: ws.write(r,c,val,num_fmt)
                            elif c==0: ws.write(r,c,val,general_fmt)
                            elif c==1: ws.write(r,c,val,text_fmt)
                            else: ws.write(r,c,val,num_fmt)
                        else:
                            if c in [3,5,8]: ws.write(r,c,val,colored_num_fmt)
                            elif 3<=c<=12 and c not in [3,5,8]: ws.write(r,c,val,num_fmt)
                            elif c==0: ws.write(r,c,val,general_fmt)
                            elif c==1: ws.write(r,c,val,text_fmt)
                            else: ws.write(r,c,val,num_fmt)

                for r in range(2,len(df)):
                    cell_value=df.iloc[r,0]
                    if isinstance(cell_value,str):
                        try:
                            hh,mm,ss=map(int,cell_value.strip().split(' ')[1].split(':'))
                            if ss!=0:
                                for c in range(df.shape[1]):
                                    val=df.iloc[r,c]
                                    if xml_type=="Den":
                                        if c==1: fmt_props={'num_format':'@','border':1,'align':'right'}
                                        elif isinstance(val,str) and val.startswith("=") or isinstance(val,(int,float)): fmt_props={'num_format':'0.00','border':1,'align':'right'}
                                        else: fmt_props={'num_format':'@','border':1,'align':'right'}
                                    elif xml_type=="Globe":
                                        if c in [3,5,10]: fmt_props={'num_format':'0.00','border':1,'align':'right','bg_color':'#B4C6E7'}
                                        elif 3<=c<=14 and c not in [3,5,10]: fmt_props={'num_format':'0.00','border':1,'align':'right'}
                                        elif c==0: fmt_props={'num_format':'General','border':1,'align':'right'}
                                        elif c==1: fmt_props={'num_format':'@','border':1,'align':'right'}
                                        else: fmt_props={'num_format':'0.00','border':1,'align':'right'}
                                    else:
                                        if c in [3,5,8]: fmt_props={'num_format':'0.00','border':1,'align':'right','bg_color':'#FFFF00'}
                                        elif 3<=c<=12 and c not in [3,5,8]: fmt_props={'num_format':'0.00','border':1,'align':'right'}
                                        elif c==0: fmt_props={'num_format':'General','border':1,'align':'right'}
                                        elif c==1: fmt_props={'num_format':'@','border':1,'align':'right'}
                                        else: fmt_props={'num_format':'0.00','border':1,'align':'right'}
                                    fmt_props['bg_color']='#FFFF00'
                                    highlight_fmt=workbook.add_format(fmt_props)
                                    if isinstance(val,str) and val.startswith("="): ws.write_formula(r,c,val,highlight_fmt)
                                    else: ws.write(r,c,val,highlight_fmt)
                        except: continue

                if xml_type=="Den": widths=[17.73,17.27]+[16.27]*7+[32.27,36.36,23.36,24.76]; hidden_cols=[6]
                elif xml_type=="Globe": widths=[17.73,17.27]+[14.91]*9+[41.91,33.27,43.36,23.36]; hidden_cols=[6,7,8]
                elif xml_type=="Glacier": widths=[17.73,17.27]+[16.91]*7+[18.73,17.55,23.36,23.36]; hidden_cols=[6]
                for i,w in enumerate(widths[:df.shape[1]]): ws.set_column(i,i,w)
                for col in hidden_cols:
                    if col<df.shape[1]: ws.set_column(col,col,None,None,{'hidden':True})

        except Exception as e:
            QMessageBox.critical(None,"Error",f"Failed to save Excel: {e}")
            if self.root: self.root.setProperty("processState","idle")
            self.thread.quit()
            self.thread.wait()
            return

        self.progressUpdated.emit(100)
        if self.root: self.root.setProperty("processState","complete")
        QMessageBox.information(None,"Done",f"Processed Excel saved:\n{save_path}")
        self.thread.quit()
        self.thread.wait()

    def resetProperties(self):
        if self.root:
            self.root.setProperty("processState","idle")
            self.root.setProperty("selectedFile","")
            self.root.setProperty("selectionType","")
            self.root.setProperty("fileSize","")
            self.root.setProperty("progress",0)
        self.selected_file=None
        self.xml_type=""
        self.progress=0

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
    sys.exit(app.exec())
