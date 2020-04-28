import xlrd
import csv
from bs4 import BeautifulSoup
import sys
import traceback
import re
import prettytable

def Output_table(list_2,flag,path_fusb,col_prov):
    str3 = ""
    str_rev = ""
    colspan =3
    if flag == False:
        str_rev = ""
        colspan = 3
    else:
        str_rev = "(Реверсивное)"
        colspan = 2

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
    return str3

def Read_Excel(filename,list_1,ind_book,col_pro,num_rows,num_col):
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
                    value_1 = str(worksheet.cell(ind_i,col_pro).value).upper()
                    value_2 = str(ind_j[1]).upper()
                    value_2 = value_2.rstrip()
                    value_2 = value_2.lstrip()
                    value_1 = value_1.rstrip()
                    value_1 = value_1.lstrip()
                    if value_1 == value_2 and value_2 != '':
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
                    value_1 = str(worksheet.cell(ind_i,col_pro).value).upper()
                    value_2 = str(ind_j[1]).upper()
                    value_2 = value_2.rstrip()
                    value_2 = value_2.lstrip()
                    value_1 = value_1.rstrip()
                    value_1 = value_1.lstrip()
                    if value_1 == value_2 or value_2 == '':
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
                    temp+= ' '
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
    soup = BeautifulSoup(html_text,'lxml')
    #Получаем список таблиц
    list_table = soup.find_all("table")
    #Проходим по всем таблицам
    for table_index in range(len(list_table)):
        soup = BeautifulSoup(str(list_table[table_index]))
        #Проходим по строкам
        list_tr = soup.find_all("tr")
        for ind_tr in range(len(list_tr)):
            soup_td = BeautifulSoup(str(list_tr[ind_tr]))
            list_td = soup_td.find_all("td")
            if ind_tr == 0:
                str1+= "*****************"+str(list_td[0].text.replace('\n',''))+"*****************\n"
            if ind_tr == 1:
                #проходим по наименованием столбцов
                th = []
                for ind_th in list_td:
                    th.append(ind_th.text.replace('\n',''))
                x = prettytable.PrettyTable(th)
                for ind_th in th:
                    x.align[ind_th] = 'l'
            if ind_tr!=0 and ind_tr!=1:
                td = []
                for ind_td in list_td:
                    td.append(ind_td.text)
                x.add_row(td)
        if len(list_tr) > 1:
            str1 += x.get_string() + "\n"
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