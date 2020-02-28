from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5 import uic
from tkinter import messagebox,Tk
import sys,os
import xlwt,xlrd,csv
def try_except(func):
        def wrapper(a):
            try:
                func(a)
            except:
                return
        return wrapper
class Startup(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon('Assets/logo.png'))
        self.setWindowTitle('Spreadsheet Application')
        self.layout=QVBoxLayout()
        self.label=QLabel()
        self.setLayout(self.layout)
        self.layout.addWidget(self.label)
        logo=QPixmap('Assets/logo.png')
        self.setAttribute(Qt.WA_TranslucentBackground )
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAcceptDrops(True)
        self.label.setPixmap(logo.scaled(300,400))
        self.show()
        self.FadeIn()
    def FadeIn(self):
        self.setWindowOpacity(0.0)
        anim=QPropertyAnimation(self,b"windowOpacity")
        anim.setDuration(3000)
        anim.setEasingCurve(QEasingCurve.OutBack)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.start(QAbstractAnimation.DeleteWhenStopped)
        anim.finished.connect(self.close)
        return QDialog.exec(self)
class SheetApp(QMainWindow):
    def __init__(self):
        super().__init__()
        start=Startup()
        uic.loadUi('MainUi.ui',self)
        self.tabWidget.setStyleSheet("QTabBar::tab { height: 50px; width: 200px;font-size: 20px; }")
        self.chosen=0
        self.show()
        self.showMaximized()
        self.sheet_choose()
        self.new=QAction(QIcon('Assets/+.png'),'New sheet')
        self.new.setToolTip('Add new module sheet')
        self.new.triggered.connect(self.sheet_choose)
        self.new.setShortcut(QKeySequence("Ctrl+n"))
        self.sizebox=QComboBox()
        self.sizebox.setToolTip('Change font size')
        self.sizebox.addItems(['8','9','10','11','12','14','16','18','20','22','24','26','28','36','48','72'])
        self.sizebox.setCurrentText('16')
        self.sizebox.currentIndexChanged.connect(self.size_change)
        self.font_size_inc=QAction(QIcon('Assets/font_size_inc.png'),'Increase')
        self.font_size_inc.setToolTip('Increase font size')
        self.font_size_inc.triggered.connect(self.font_size_increase)
        self.font_size_dec=QAction(QIcon('Assets/font_size_dec.png'),'Decrease')
        self.font_size_dec.setToolTip('Decrease font size')
        self.font_size_dec.triggered.connect(self.font_size_decrease)
        self.delete_row=QAction(QIcon('Assets/delete_row.png'),'Delete Row')
        self.delete_row.setToolTip('Delete Currently selected row')
        self.delete_row.triggered.connect(self.delete_current_row)
        self.insert_row1=QAction(QIcon('Assets/insert_row_before.png'),'Insert before')
        self.insert_row1.setToolTip('Insert row before current')
        self.insert_row1.triggered.connect(self.add_before)
        self.insert_row2=QAction(QIcon('Assets/insert_row_after.png'),'Insert after')
        self.insert_row2.setToolTip('Insert row after current')
        self.insert_row2.triggered.connect(self.add_after)
        self.toolBar.addAction(self.new)
        self.toolBar.addWidget(self.sizebox)
        self.toolBar.addAction(self.font_size_inc)
        self.toolBar.addAction(self.font_size_dec)
        self.toolBar.addAction(self.delete_row)
        self.toolBar.addAction(self.insert_row1)
        self.toolBar.addAction(self.insert_row2)
        self.toolBar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.tabWidget.tabCloseRequested.connect(self.removeTab)
        self.setStyleSheet('''
        QPushButton{font-size: 20px;}
        QHeaderView
        {
        border: 1px solid grey;
        text-align: centre;
        font-family: arial;
        font-size:20px;
        }
        ''')
    @try_except
    def delete_current_row(self):
        tables = self.tabWidget.currentWidget().findChildren(QTableWidget)
        tables[0].removeRow(tables[0].currentRow())
    @try_except
    def add_before(self):
        tables = self.tabWidget.currentWidget().findChildren(QTableWidget)
        tables[0].insertRow(tables[0].currentRow())
    @try_except
    def add_after(self):
        tables = self.tabWidget.currentWidget().findChildren(QTableWidget)
        tables[0].insertRow(tables[0].currentRow()+1)
    @try_except
    def size_change(self):
        tables = self.findChildren(QTableWidget)
        for table in tables:
            font=table.font()
            font.setPointSize(int(self.sizebox.currentText()))
            table.setFont(font)
            table.resizeColumnsToContents()
            table.resizeRowsToContents()
    @try_except
    def font_size_increase(self):
        tables = self.findChildren(QTableWidget)
        for table in tables:
            font=table.font()
            font.setPointSize(font.pointSize()+1)
            table.setFont(font)
            table.resizeColumnsToContents()
            table.resizeRowsToContents()
    @try_except
    def font_size_decrease(self):
        tables = self.findChildren(QTableWidget)
        for table in tables:
            font=table.font()
            font.setPointSize(font.pointSize()-1)
            table.setFont(font)
            table.resizeColumnsToContents()
            table.resizeRowsToContents()
    def closeEvent(self,event):
        Tk().wm_withdraw()
        msg=messagebox.askokcancel('Exit','Warning:Are you sure, you want to exit?\n All your unsaved changes will be lost if you continue')
        if(msg==1):
            event.accept()
        else:
            event.ignore()
    def removeTab(self, index):
        Tk().wm_withdraw()
        msg=messagebox.askokcancel('Closing Tab','Warning: All your unsaved changes will be lost if you continue')
        if msg==1:
            widget = self.tabWidget.widget(index)
            if widget != None:
                widget.deleteLater()
                self.tabWidget.removeTab(index)
    def sheet_choose(self):
        ch=Choice(self)
        ch.exec()
        Home=Sheet(self,self.chosen)
        self.chosen=0
class Choice(QDialog):
    def __init__(self,ui_var):
        super().__init__()
        uic.loadUi('Choice.ui',self)
        self.setWindowTitle('Choose Sheet Type')
        self.setWindowIcon(QIcon('Assets/logo.png'))
        self.obj=ui_var
        self.pushButton.clicked.connect(self.option1)
        self.pushButton_2.clicked.connect(self.option2)
        self.pushButton_3.clicked.connect(self.option3)
        self.pushButton_4.clicked.connect(self.option4)
    def option1(self):
        self.obj.chosen=1
        self.close()
    def option2(self):
        self.obj.chosen=2
        self.close()
    def option3(self):
        self.obj.chosen=3
        self.close()
    def option4(self):
        self.obj.chosen=4
        self.close()
class Sheet(QWidget):
    def __init__(self,ui_var,typ):
        super().__init__()
        uic.loadUi('Sheet.ui', self)
        self.typ=typ
        self.obj=ui_var
        self.cells_add=QAction('Add Cells',self)
        self.cells_add.triggered.connect(self.add_rows)
        self.cells_add.setShortcut(Qt.Key_Return)
        self.addAction(self.cells_add)
        if(self.typ==0):
            return         
        p=QPixmap('Assets/info.png')
        self.label.setPixmap(p.scaled(60,60))
        self.modules={1:'FinPlate',2:'TensionMember',3:'BCEndPlate',4:'CleatAngle'}
        tooltips={
            1:'''
            <b>Note:</b><br>
            <table border=2>
            <tr>
            <th>Connection Type
            <th>Description
            </tr>
            <tr>
            <td>1
            <td>Column flange - Beam web
            </tr>
            <tr>
            <td>2
            <td>Column web - Beam web
            </tr>
            <tr>
            <td>3
            <td>Beam - Beam
            </tr>
            </table>
        ''',
        2:'''
            <b>Note:</b><br>
            <table border=2>
            <tr>
            <th>Support conditions
            <th>Description
            </tr>
            <tr>
            <td>1
            <td>Fixed
            </tr>
            <tr>
            <td>2
            <td>Hinged
            </tr>
            </table>
        ''',
        3:'''
            <b>Note:</b><br>
            <table border=2>
            <tr>
            <th>End Plate type
            <th>Description
            </tr>
            <tr>
            <td>1
            <td>Extended One Way
            </tr>
            <tr>
            <td>2
            <td>Extended Both Ways
            </tr>
            <tr>
            <td>3
            <td>Flush End Plate
            </tr>
            </table>
        ''',
        4:'''
        <table border=2>
        <tr>
        <th>No Description available for this sheet
        </tr>
        </table>
        '''}
        self.label.setToolTip(tooltips[self.typ])
        if(self.typ==1):
            self.headers=['ID','Connection type','Axial load','Shear load','Bolt diameter','Bolt grade','Plate thickness']
        elif(self.typ==2):
            self.headers=['ID','Member length','Tensile load','Support condition at End 1','Support condition at End 2']
        elif(self.typ==3):
            self.headers=['ID','End plate type','Shear load','Axial Load','Moment Load','Bolt diameter','Bolt grade','Plate thickness']
        elif(self.typ==4):
            self.headers=['ID','Angle leg 1','Angle leg 2','Angle thickness','Shear load','Bolt diameter','Bolt grade']
        for i in range(len(self.headers)):
            self.tableWidget.insertColumn(i)
        self.tableWidget.setHorizontalHeaderLabels(self.headers)
        self.obj.tabWidget.addTab(self,self.modules[self.typ])
        self.tableWidget.setRowCount(100)
        self.setStyleSheet('''
        QPushButton{font-size: 14;}
        QHeaderView
        {
        border: 1px solid grey;
        text-align: centre;
        font-family: arial;
        font-size:20px;
        }
        ''')
        font=QFont()
        font.setPointSize(15)
        self.tableWidget.setFont(font)
        self.tableWidget.resizeColumnsToContents()
        self.pushButton_3.clicked.connect(self.create_sample_file)
        self.pushButton.clicked.connect(self.load_input)
        self.pushButton_4.clicked.connect(self.validate)
        self.pushButton_2.clicked.connect(self.submit)
        self.pushButton_5.clicked.connect(self.reset)
    def reset(self):
        Tk().wm_withdraw()
        msg=messagebox.askokcancel('Reset','Warning: All your unsaved changes will be lost if you continue')
        if(msg==1):
            self.tableWidget.setRowCount(0)
            self.tableWidget.setRowCount(100)
    def submit(self):
        if(self.validate(True)):
            count=True
            fileName = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
            if(fileName==''):
                Tk().wm_withdraw()
                messagebox.showerror('Error','No folder chosen.')
                return
            for i in range(self.tableWidget.rowCount()):
                my_dict={}
                for j in range(self.tableWidget.columnCount()):
                    if(self.tableWidget.item(i,j)==None or self.tableWidget.item(i,j).text()==""):
                        my_dict[self.headers[j]]=None
                    else:
                        if(j==0):
                            my_dict[self.headers[j]]=int(self.tableWidget.item(i,j).text())
                        else:
                            my_dict[self.headers[j]]=float(self.tableWidget.item(i,j).text())
                if all(ele == None for ele in my_dict.values()):
                    continue
                else:
                    count=False
                    f=open(fileName+'/'+self.modules[self.typ]+'_'+str(my_dict['ID'])+'.txt','w')
                    f.write(str(my_dict))
                    f.close()
            if(count):
                Tk().wm_withdraw()
                messagebox.showerror('Error','The spreadsheet is empty.\nFill in some values to submit it properly.')
            else:
                Tk().wm_withdraw()
                messagebox.showinfo('Success',f'Sheet Submitted successfully and compiled into multiple files\nCheck {fileName} for your file')
    def validate(self,check=False):
        errors=[]
        dupli_check=[]
        count=0
        for i in range(self.tableWidget.rowCount()):
            try:
                if(type(eval(self.tableWidget.item(i,0).text()))==int):
                    dupli_check.append(self.tableWidget.item(i,0).text())
            except:
                pass
            for j in range(self.tableWidget.columnCount()):
                try:
                    if(self.tableWidget.item(i,j)==None):
                        continue
                    elif(self.tableWidget.item(i,j).text()==""):
                        if(self.tableWidget.item(i,j).background().color()==Qt.red):
                            self.tableWidget.item(i,j).setBackground(Qt.white)
                        continue
                    else:
                        if(self.tableWidget.item(i,j).background().color()==Qt.red):
                            self.tableWidget.item(i,j).setBackground(Qt.white)
                        float(self.tableWidget.item(i,j).text())
                except:
                    try:
                        errors.append((i+1,j+1,self.tableWidget.item(i,j).text(),'NON-NUMERICAL'))
                        count+=1
                    except:
                        pass
        dupli_indices=[idx for idx, val in enumerate(dupli_check) if val in dupli_check[:idx]]
        for i in dupli_indices:
            try:
                errors.append((i+1,1,dupli_check[i],'DUPLICATE'))
                count+=1
            except:
                pass
        if(count>0):
            err=ErrorDialog(sorted(errors),self)
            err.exec()
            if(check):
                return(False)
        else:
            if(check):
                return(True)
            else:
                Tk().wm_withdraw()
                messagebox.showinfo('Validation','Checked:\nNo errors found.\nAll records are correctly entered.\nYou may proceed and submit now')

    def add_rows(self):
        if(self.tableWidget.currentRow()>=self.tableWidget.rowCount()-10):
            self.tableWidget.setRowCount(self.tableWidget.rowCount()+10)
    def load_input(self):
        Tk().wm_withdraw()
        messagebox.showwarning('Importing','Warning: All your unsaved changes will be lost if you continue')
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Spreadsheet (*.csv *.xls *.xlsx)", options=options)
        if(fileName==''):
                Tk().wm_withdraw()
                messagebox.showerror('Error','No file chosen.')
                return
        try:
            if(fileName[-3:]=='csv'):
                with open(fileName) as csv_file:
                    sheet = list(csv.reader(csv_file, delimiter=','))
                    self.tableWidget.setRowCount(len(sheet)+9)
                    for i in range(1,len(sheet)):
                        for j in range(len(self.headers)):
                            if(j==0):
                                self.tableWidget.setItem(i-1 , j,QTableWidgetItem(str(int(sheet[i][j]))))
                            else:
                                self.tableWidget.setItem(i-1 , j,QTableWidgetItem(str(sheet[i][j])))
                    self.tableWidget.resizeColumnsToContents()
            else:
                wb = xlrd.open_workbook(fileName)
                sheet = wb.sheet_by_index(0)
                self.tableWidget.setRowCount(sheet.nrows+9)
                for i in range(1,sheet.nrows):
                    for j in range(len(self.headers)):
                        if(j==0):
                            self.tableWidget.setItem(i-1 , j,QTableWidgetItem(str(int(sheet.cell_value(i, j)))))
                        else:
                            self.tableWidget.setItem(i-1 , j,QTableWidgetItem(str(sheet.cell_value(i, j))))
                self.tableWidget.resizeColumnsToContents()
        except Exception as e:
            Tk().wm_withdraw()
            messagebox.showerror('Error',f'File not readable.\nError: {e}')
    def create_sample_file(self):
        wb=xlwt.Workbook()
        style = xlwt.easyxf('font: bold 1')
        sheet1 = wb.add_sheet(self.modules[self.typ])
        for i in range(len(self.headers)):
            sheet1.write(0, i, self.headers[i],style)
        wb.save(os.environ['USERPROFILE']+'\\desktop\\'+'Sample_'+self.modules[self.typ]+'.xls')
        Tk().wm_withdraw()
        messagebox.showinfo('File Created','Sample File successfully created on the desktop !')
class ErrorDialog(QDialog):
    def __init__(self,errors,ui_var):
        super().__init__()
        self.obj=ui_var
        uic.loadUi('Error.ui', self)
        self.setWindowTitle('Errors')
        self.setWindowIcon(QIcon('Assets/error.png'))
        self.label.setText('Error Count: '+str(len(errors)))
        self.tableWidget.setRowCount(len(errors))
        self.pushButton.clicked.connect(self.close)
        for i in range(len(errors)):
            self.tableWidget.setItem(i,0,QTableWidgetItem(str(errors[i][0])))
            self.tableWidget.setItem(i,1,QTableWidgetItem(str(errors[i][1])))
            self.tableWidget.setItem(i,2,QTableWidgetItem(str(errors[i][2])))
            self.tableWidget.setItem(i,3,QTableWidgetItem(str(errors[i][3])))
            self.obj.tableWidget.item(errors[i][0]-1,errors[i][1]-1).setBackground(Qt.red)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.setStyleSheet('''
        QTableWidget{
            background-color: #EB7775;
        }
        QHeaderView
        {
        color: #EB7775;
        border: 1px solid grey;
        text-align: centre;
        }
        ''')
app=QApplication(sys.argv)
window=SheetApp()
app.exec_()