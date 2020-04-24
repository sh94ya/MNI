import traceback
import xlrd
import sys  # sys нужен для передачи argv в QApplication
from PyQt5 import QtWidgets
import design  # Это наш конвертированный файл дизайна
import os
import function_read as ff
from bs4 import BeautifulSoup

class Usage(Exception):
    def __init__(self, msg):
        self.msg = msg

class ExampleApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    global path_excel
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton.clicked.connect(self.browse_folder_excel)
        self.pushButton_2.clicked.connect(self.browse_folder_fusb)
        self.pushButton_3.clicked.connect(self.start)
        self.checkBox.clicked.connect(self.checkBox_clicked)
        self.spinBox.valueChanged.connect(self.spinBox_valchange)
        self.pushButton_4.clicked.connect(self.clear)
        self.pushButton_5.clicked.connect(self.safe_rez)
        self.checkBox_2.clicked.connect(self.checkBox_2clicked)

    def browse_folder_excel(self):
        global path_excel
        fname = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file','',"Excel files (*.xls *.xlsx)")
        # открыть диалог выбора директории и установить значение переменной
        # равной пути к выбранной директории
        if fname:  # не продолжать выполнение, если пользователь не выбрал директорию
            path_excel = fname[0]
            self.lineEdit.setText(path_excel)
            try:
                workbook = xlrd.open_workbook(path_excel)
                worksheet = workbook.sheet_by_index(0)
                nr = str(worksheet.nrows)
                nc = str(worksheet.ncols)
                self.spinBox_2.setValue(worksheet.nrows)
                self.spinBox_3.setValue(worksheet.ncols)
            except BaseException:
                print('Ошибка чтения файла')

    def browse_folder_fusb(self):
        global path_fusb
        fname = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file','',"FUSB files (*.csv *.html *.txt)")
        # открыть диалог выбора директории и установить значение переменной
        # равной пути к выбранной директории
        if fname:  # не продолжать выполнение, если пользователь не выбрал директорию
            path_fusb = fname[0]
            self.lineEdit_3.setText(path_fusb)
    
    def checkBox_clicked(self):
        if self.checkBox.checkState() != 0:
            self.spinBox_2.setEnabled(False)
            self.spinBox_3.setEnabled(False)
        else:
            self.spinBox_2.setEnabled(True)
            self.spinBox_3.setEnabled(True)
    
    def checkBox_2clicked(self):
        global flag
        if self.checkBox_2.checkState() != 0:
            flag = True
        else:
            flag = False


    def spinBox_valchange(self):
        global index_book
        index_book = self.spinBox.value()

    def start(self):
        global list_1,index_book,col_prov,num_rows,num_col,flag
        try:
            if path_fusb!='' and path_excel!='':
                filename, file_extension = os.path.splitext(path_fusb)
                if file_extension == '.csv':
                    list_1 = ff.Read_FUSB_CSV(path_fusb)
                if file_extension == '.html':
                    list_1 = ff.Read_FUSB_HTML(path_fusb)
                if file_extension == '.txt':
                    list_1 = ff.Read_FUSB_TXT(path_fusb)
                index_book = self.spinBox.value()                
                num_rows = self.spinBox_2.value()
                num_col = self.spinBox_3.value()
                col_prov = self.spinBox_4.value()-1
                str_rev = ""
                colspan = 3
                if flag == False:
                    list_2 = ff.Read_Excel(path_excel,list_1,index_book,col_prov,num_rows,num_col)
                    str_rev = ""
                    colspan = 3
                else:
                    list_2 = ff.Read_Excel_reverse(path_excel,list_1,index_book,col_prov,num_rows,num_col)  
                    str_rev = "(Реверсивное отображение)"  
                    colspan = 2                
                str3 = ""
                if len(list_2)>0:
                    str3 += ('<table cellspacing="2" border="1" cellpadding="5"><tr><td align="center" colspan="'+str(colspan)+'"><b>'+str_rev+" Найдены совпадения для "+path_fusb+'</b></td></tr><thead><tr bgcolor="#8a7f8e"><th>№</th>')
                    if flag == False:
                        str3 += "<th>Excel</th>"
                    str3 += "<th>FUSB</th></tr></thead>"
                    count = 0
                    for row_list in list_2:
                        str3+=('<tr>')
                        str3+=('<td width="20%"><pre>'+str(count)+'</pre></td>')
                        ccount = 0
                        if flag == False:
                            str3+=("<td><pre>") 
                            #str3+= '<table width="100%"><tr>'                            
                            for item_row in row_list[0]:
                                if ccount == col_prov:
                                        str3+=('<font size="4"><b>'+str(item_row)+"|</b></font>")
                                else:
                                        str3+=(''+str(item_row)+"|")
                                ccount += 1
                           # str3+= "</tr></table>"                                  
                            str3+=("</pre></td>") 
                        str3+=("<td><pre>") 
                        ccount = 0
                        for item_row in row_list[1]:
                            if ccount == 1:
                                 str3+=('<font size="4"><b>'+str(item_row)+"|</b></font>")
                            else:
                                 str3+=(""+str(item_row)+" |")
                            ccount += 1
                        str3+=("</pre></td>") 
                        str3+=("</tr>")
                        count+=1
                    str3+=("</table><br>")
                else:
                      str3+=('<table cellspacing="2" border="1" cellpadding="5"><tr><td align="center" colspan="3"><b>'+str_rev+" Cовпадений для "+path_fusb+' не найдено</b></td></tr></table><br>')
                self.textEdit.append(str3)
                #self.textEdit.append("<hr>")
        except Exception as e:
            print('Ошибка:\n', traceback.format_exc())

    def clear(self):
        self.textEdit.setText("")
    
    def safe_rez(self):
        global bs4
        fname = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file','',"TXT files (*.txt);;HTML files (*.html)")
        if fname[0]:
            f = open(fname[0], 'w')
            with f:
                filename, file_extension = os.path.splitext(fname[0])
                if file_extension == '.txt':
                    str1 = ff.output_html_to_txt(self.textEdit.toHtml())
                    f.write(str1)
                if file_extension == '.html':
                    f.write(self.textEdit.toHtml())
                
####varible####                
flag = False
list_1 = []
path_fusb = ''
path_excel = ''
index_book = 1
num_rows = 0
num_col = 0
col_prov = 1
header_fusb = []
header_excel = []
###############  
def main():
    list_1 = []
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

if __name__ == "__main__":
    main()