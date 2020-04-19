import xlrd
import csv
import getopt
from bs4 import BeautifulSoup
import sys
import traceback
import re
import pandas as pd
import prettytable

def Read_Excel(filename,list_1,ind_book,col_pro,num_rows,num_col):
    st = ""
    line_1 = []
    try:
        workbook = xlrd.open_workbook(filename)
        worksheet = workbook.sheet_by_index(ind_book-1)
    except BaseException:
        print('Ошибка чтения файла')
        return
    try:
        ind_i = 0
        count = 0
        for ind_i in range(num_rows):
            for ind_j in list_1:
                try:
                    if worksheet.cell(ind_i,col_pro).value == ind_j[1] and  ind_j[1] != '':
                        count+=1
                        line_1.append([worksheet.row_values(ind_i),ind_j])
                        break
                except Exception as e:
                    print('Ошибка:\n', traceback.format_exc())
            ind_i+=1
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return line_1

def Read_Excel_reverse(filename,list_1,ind_book,col_pro,num_rows,num_col):
    st = ""
    line_1 = []
    try:
        workbook = xlrd.open_workbook(filename)
        worksheet = workbook.sheet_by_index(ind_book-1)
    except BaseException:
        print('Ошибка чтения файла')
        return
    try:
        ind_i = 0
        for ind_j in list_1:  
            count = 0  
            for ind_i in range(num_rows):
                try:
                    if str(worksheet.cell(ind_i,col_pro).value) == ind_j[1]:
                        count+=1
                        break
                    if  str(ind_j[1]) == '':
                        count+=1
                        break
                except Exception as e:
                    print('Ошибка:\n', traceback.format_exc())
                    count+=1
                ind_i+=1
            if count == 0:
                line_1.append([[],ind_j])
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return line_1  

def listToString(s):  
    str1 = ""  
    for ele in s:  
        str1 += str(ele) + "  "  
    str1+= ""
    return str1  

#Чтение файла fusb в формате txt
def Read_FUSB_TXT(filename):
    list_1 = []
    try:
        list_1 = []
        with open(filename, 'r') as file:
            reader = csv.reader(file,delimiter = '|')
            for row in reader:
                try:
                    list_2 = []
                    for cell in row:
                        list_2.append(re.sub(r'\s+', '', cell))
                    del list_2[0]
                    if len(list_2)>0:
                        list_1.append(list_2)
                except Exception as e:
                    print('Ошибка:\n', traceback.format_exc())
            parse_txt_list(list_1)
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return list_1

#Чтение файла fusb в формате csv
def Read_FUSB_CSV(filename):
    list_1 = []
    try:
        list_1 = []
        with open(filename, 'r') as file:
            reader = csv.reader(file,delimiter = ';')
            for row in reader:
                list_2 = []
                for cell in row:
                    list_2.append(cell)
                list_1.append(list_2)
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return list_1

#Чтение файла fusb в формате html
def Read_FUSB_HTML(filename):
    list_1 = []
    list_tr = []
    try:
        with open(filename, "r") as f:
            contents = f.read()
            soup = BeautifulSoup(contents, 'lxml')
        ##Выводим все строки таблицы в список
        for tag in soup.find_all("tr"):
            list_tr.append(tag)
        ##Парсим полученый выше список
        ind_i = 0
        for item in list_tr:
            soup1 = BeautifulSoup(str(item))
            temp = []
            for tag in soup1.find_all(['td', 'th']):
                temp.append(tag.text)
            list_tr[ind_i] = temp
            ind_i+=1  
        list_1 = parse_html_list(list_tr)
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return list_1

##Преобразуем предыдущий распарсенный список для html-документов
def parse_html_list(list):
    list_1 = []
    for ind_i in range(len(list)):
        if len(list[ind_i]) == 1:
            ind_j = ind_i+1
            while ind_j<len(list) and len(list[ind_j]) != 1:
                ind_k = 0
                try:
                    for ind_k in range(len(list[ind_j])):
                        if list[ind_j][ind_k] == '\n\n':
                            del list[ind_j][ind_k]
                            ind_k-=1
                            break
                        else:
                            del list[ind_j][ind_k]
                            ind_k-=1
                        ind_k+=1
                except Exception as e:
                    print('Ошибка:\n', traceback.format_exc())
                temp = list[ind_j]
                temp[0] = ((list[ind_i][0])+" "+ temp[0])
                list_1.append(temp)
                ind_j+=1
        ind_i+=1
    return list_1

def parse_txt_list(list):
    list_1 = []
    try:
        del list[0] #удаляем первый элемент с названиями стобцов
        ind_i = 0
        for item in list:
            if len(item) == 2:
                ind_j = list.index(item)+1
                try:
                    while len(list[ind_j]) != 2:
                        list[ind_j][0] = item[0] + " " + list[ind_j][0]
                        ind_j += 1
                except Exception as e:
                    print('Ошибка:\n', traceback.format_exc())
                del list[list.index(item)]
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return list_1

def output_html_to_txt(html_text):
    str1 = ""
    list_header = []
    list_table = []
    df = pd.read_html(html_text)
    try:
        for item in df:
            header = list(item.columns.values)
            tbody = item.values.tolist()
            x = prettytable.PrettyTable(header)
            for item_list in tbody:
                x.add_row(item_list)
            str1 += x.get_string() + "\n"
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return str1