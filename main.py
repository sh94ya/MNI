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
        self.pushButton_4.clicked.connect(self.clear)
        self.pushButton_5.clicked.connect(self.safe_rez)
        self.checkBox_2.clicked.connect(self.checkBox_2clicked)

    def browse_folder_excel(self):
        global path_excel,num_rows,num_col
        fname = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file','',"Excel files (*.xls *.xlsx)")
        # открыть диалог выбора директории и установить значение переменной
        # равной пути к выбранной директории
        if fname:  # не продолжать выполнение, если пользователь не выбрал директорию
            path_excel = fname[0]
            self.lineEdit.setText(path_excel)
            try:
                workbook = xlrd.open_workbook(path_excel)
                worksheet = workbook.sheet_by_index(0)
                num_rows = (worksheet.nrows)
                num_col = (worksheet.ncols)
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
    
   
    def checkBox_2clicked(self):
        global flag
        if self.checkBox_2.checkState() != 0:
            flag = True
        else:
            flag = False

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
                str_rev = ""
                if flag == False:
                    list_2 = ff.Read_Excel(path_excel,list_1,index_book,col_prov,num_rows,num_col)
                    str_rev = ""
                else:
                    list_2 = ff.Read_Excel_reverse(path_excel,list_1,index_book,col_prov,num_rows,num_col)  
                    str_rev = "(Реверсивное отображение)"                  
                str3 = ""
                if len(list_2)>0:
                    str3 += ('<div name="header"><p  name="header"><h1>'+str_rev+" Найдены совпадения для "+path_fusb+'</h1></p></div><div name="table"><table title="Найдены совпадения для '+path_fusb+ 'cellspacing="2" border="1" cellpadding="5"><thead><tr><th>№</th>')
                    if flag == False:
                        str3 += "<th>Excel</th>"
                    str3 += "<th>FUSB</th></tr></thead>"
                    count = 1
                    for row_list in list_2:
                        str3+=("<tr>") 
                        str3+=("<td>"+str(count)+"</td>")
                        ccount = 0
                        if flag == False:
                            str3+=("<td>") 
                            for item_row in row_list[0]:
                                if ccount == col_prov-1:
                                        str3+=("<b>"+str(item_row)+"</b>")
                                else:
                                        str3+=("| "+str(item_row)+" |")
                                ccount += 1
                            str3+=("|</td>") 
                        str3+=("<td>") 
                        ccount = 0
                        for item_row in row_list[1]:
                            if ccount == 1:
                                 str3+=("<b>"+str(item_row)+"</b>")
                            else:
                                 str3+=("| "+str(item_row)+" |")
                            ccount += 1
                        str3+=("|</td>") 
                        str3+=("</tr>")
                        count+=1
                    str3+=("</table></div>")
                else:
                      str3+=('<div name="header"><p align="left"><h1>'+str_rev+" Совпадений для "+path_fusb+' не найдено</h1></p></div><div name="table"></div>')
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
                
flag = False
list_1 = []
path_fusb = ''
path_excel = ''
index_book = 1
num_rows = 0
num_col = 0
col_prov = 1
def main():
    list_1 = []
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

if __name__ == "__main__":
    main()