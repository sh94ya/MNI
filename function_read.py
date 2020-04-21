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
    line_ex = []
    line_f = []
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
                        #line_1.append([worksheet.row_values(ind_i),ind_j])
                        line_ex.append(worksheet.row_values(ind_i))
                        line_f.append(ind_j)
                except Exception as e:
                    print('Ошибка:\n', traceback.format_exc())
            ind_i+=1
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    try:
        if len(line_f) > 0:
            line_f.insert(0,list_1[0])
            line_ex.insert(0,worksheet.row_values(0))            
            normalize_view_data(line_ex)
            normalize_view_data(line_f)
        for i in range(len(line_ex)):
            line_1.append([line_ex[i],line_f[i]])
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return line_1

def Read_Excel_reverse(filename,list_1,ind_book,col_pro,num_rows,num_col):
    st = ""
    line_1 = []
    line_ex = []
    line_f = []
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
                line_f.append(ind_j)
                line_1.append([[],ind_j])
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    try:
        if len(line_f) > 0:
            line_f.insert(0,list_1[0])       
            normalize_view_data(line_f)
        for i in range(len(line_ex)):
            line_1.append([[],line_f[i]])
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return line_1  

def normalize_view_data(list):
    list_1 = []
    for item in list[0]:
        list_1.append(0)
    #заполняем массив с максимальными размерами строк
    for item in list:
        for ob in item:
            if len(str(ob)) > list_1[item.index(ob)]:
                list_1[item.index(ob)] = len(str(ob))
    #дополняем пробелами
    for i in range(len(list)):
        for j in range(len(list[i])):
            if len(str(list[i][j])) < list_1[j]:
                temp = ''
                for index in range(list_1[j] - len(str(list[i][j]))):
                    temp+= '_'
                list[i][j] = str(list[i][j]) + temp + ""

def listToString(s):  
    str1 = ""  
    for ele in s:  
        str1 += str(ele) + "  "  
    str1+= ""
    return str1  

#Чтение файла fusb в формате txt
def Read_FUSB_TXT(filename):
    global header_fusb
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
        try:
            header_fusb = list_1[0]
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
        del list_1[0]
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
    list_1.append(list[0])
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
        #del list[0] #удаляем первый элемент с названиями стобцов
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
            str1+= "*****************"+list(item.columns.values)[0]+"*****************\n"
            tbody = item.values.tolist()
            x = prettytable.PrettyTable(tbody[0])
            #равнение столбцов
            for th in tbody[0]:
                x.align[th] = 'l'
            for item_list in tbody:
                if tbody.index(item_list)!=0:
                    x.add_row(item_list)
            str1 += x.get_string() + "\n"
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return str1

def output_to_txt(list,header):
    str1 = ""
    list_header = []
    list_table = []
    try:
        tbody = list
        x = prettytable.PrettyTable(header)
        #равнение столбцов
        for th in header:
            x.align[th] = 'l'
        for item_list in tbody:
            x.add_row([listToString(item_list[0]),listToString(item_list[1])])
        str1 += x.get_string() + "\n"
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return str1