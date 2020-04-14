import xlrd
import csv
import getopt
from bs4 import BeautifulSoup
import sys
import traceback
import re

def Read_Excel(filename,list_1,ind_book,col_pro,num_rows,num_col):
    st = ""
    try:
        workbook = xlrd.open_workbook(filename)
        worksheet = workbook.sheet_by_index(ind_book-1)
        #num_rows=worksheet.nrows
        #num_col=worksheet.ncols
    except BaseException:
        print('Ошибка чтения файла')
        return
    try:
        ind_i = 0
        count = 0
        for ind_i in range(num_rows):
            for ind_j in list_1:
                try:
                    if worksheet.cell(ind_i,col_pro-1).value == ind_j[1] and  ind_j[1] != '':
                        count+=1
                        #str2 = listToString(ind_j)
                        st += str(count)+") Excel_OUT #"+listToString(worksheet.row_values(ind_i))+"# - FUSB_OUT #"+listToString(ind_j)+"#\n"
                        #print(count,") ",worksheet.row_values(ind_i),"-",ind_j)
                        break
                except Exception as e:
                    print('Ошибка:\n', traceback.format_exc())
            ind_i+=1
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
    return st

def listToString(s):  
    str1 = "["  
    for ele in s:  
        str1 += str(ele) + " | "  
    str1+= "]"
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

##Преобразуем предыдущий распарсенный список
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
                print(temp)
                list_1.append(temp)
                ind_j+=1
        ind_i+=1
    return list_1