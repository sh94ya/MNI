import traceback
import xlrd
import sys  # sys нужен для передачи argv в QApplication
from PyQt5 import QtWidgets
import design  # Это наш конвертированный файл дизайна
import os
import function_read as ff

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
    
    def spinBox_valchange(self):
        global index_book
        index_book = self.spinBox.value()

    def start(self):
        global list_1,index_book,col_prov,num_rows,num_col
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
                col_prov = self.spinBox_4.value()
                self.textEdit.setText(self.textEdit.toPlainText() + "\n*****************\n"+ "БД МНИ: "+path_excel+"   FUSB: "+path_fusb+"\n")
                str = ff.Read_Excel(path_excel,list_1,index_book,col_prov,num_rows,num_col)
                self.textEdit.setText(str)
                self.textEdit.setText(self.textEdit.toPlainText() + "\n*****************\n")
        except Exception as e:
            print('Ошибка:\n', traceback.format_exc())

    def clear(self):
        self.textEdit.setText("")

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