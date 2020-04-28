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
        # self.checkBox_2.clicked.connect(self.checkBox_2clicked)

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
    
   
    # def checkBox_2clicked(self):
    #     global flag
    #     if self.checkBox_2.checkState() != 0:
    #         flag = True
    #     else:
    #         flag = False

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
                #####Вывод в text_edit########        
                list_2 = ff.Read_Excel(path_excel,list_1,index_book,col_prov,num_rows,num_col)       
                self.textEdit.append(ff.Output_table(list_2,False,path_fusb,col_prov))
                list_3 = ff.Read_Excel_reverse(path_excel,list_1,index_book,col_prov,num_rows,num_col)
                self.textEdit_2.append(ff.Output_table(list_3,True,path_fusb,col_prov))
        except Exception as e:
            print('Ошибка:\n', traceback.format_exc())

    def clear(self):
        self.textEdit.setText("")
        self.textEdit_2.setText("")
    
    def safe_rez(self):
        global bs4
        fname = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file','',"TXT files (*.txt);;HTML files (*.html)")
        if fname[0]:
            filename, file_extension = os.path.splitext(fname[0])
            f = open(filename+'(Совпадающие c БД)'+file_extension, 'w')
            with f:
                if file_extension == '.txt':
                    str1 = ff.output_html_to_txt(self.textEdit.toHtml())
                    f.write(str1)
                if file_extension == '.html':
                    f.write(self.textEdit.toHtml())
            f = open(filename+'(Несовпадающие c БД)'+file_extension, 'w')
            with f:
                if file_extension == '.txt':
                    str1 = ff.output_html_to_txt(self.textEdit_2.toHtml())
                    f.write(str1)
                if file_extension == '.html':
                    f.write(self.textEdit_2.toHtml())

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