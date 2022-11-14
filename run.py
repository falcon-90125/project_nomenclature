import pandas as pd
import datetime
#ЕСЛИ НОВЫЙ ПРАЙС-ЛИСТ ШАПКУ ПРАЙС-ЛИСТА СКОПИРОВАТЬ ИЗ ПРЕДЫДУЩЕЙ ВЕРСИИ
'''ВВЕСТИ ДИРЕКТОРИЮ ДЛЯ СОХРАНЕНИЯ ФАЙЛА'''
file_directory_resalts = 'E:\Соколов Алексей\Desktop/'

todays_date = datetime.date.today() #сегодняшняя дата

file_name = 'Номенклатура_CТР_' + str(todays_date) + '.xlsx' #имя файла

price_varton = pd.read_csv('E:\VSCode\Nomenklatura_STR_VARTON/price_varton.csv', encoding='cp1251', delimiter=';', index_col = 0) #загружаем прайс

f = open('E:\Соколов Алексей\Desktop/Номенклатура_СТР.txt', 'r', encoding='utf-8') #загружаем файл номенклатуры СТР
nom_load = f.read()
f.close()
print(nom_load)
nom_load =  nom_load[nom_load.index('Место выхода света')+18:] #Отбрасываем шапку

vendor_code_list = [] # Список с артикулами
numbers = [] # Список с количествами

for i in range(nom_load.count('Комплектация')): # Проходим по тексту столько раз, сколько стречается слово 'Комплектация' - это соответствует кол-ву артикулов
    #Выцепляем кол-во позиции и добавляем в список
    text_snippet = nom_load[nom_load.index(' - ')-27:nom_load.index(' - ')] # Фрагмент текста от (Место выхода света) до ' - ' перед артикулом
    kol_vo =''
    for i in range(len(text_snippet)): # Проходим по фрагменту и проверяем наличие чисел по списку chislo_list
        chislo_list = ['0','1','2','3','4','5','6','7','8','9']
        if text_snippet[i] in chislo_list: # Если есть число добавляем в переменную kol_vo
            kol_vo = kol_vo + text_snippet[i]
    numbers.append(int(kol_vo))  # Кол-во добавляем в список

    #Выцепляем артикул и остальную инфу до 'Место выхода света'
    art_predv = nom_load[nom_load.index(' - ')+3 : nom_load.index('Светоотдача')]

    if '+' in art_predv: # Если два артикула через +, например аварийник с пинтограммой
        # Выцепляем первый артикул
        art_predv_1 = art_predv[:art_predv.index(' ')]
        # Избавляемся от первого артикула и выцепляем всё что есть после +
        art_predv_2_1 = art_predv[art_predv.index('+')+2 : art_predv.index('+')+90].replace('\n','')
        # Выцепляем перваю часть второго артикула
        art_predv_2_2 = art_predv_2_1[:art_predv_2_1.index('  ')]
        # Выцепляем вторую часть артикула, но необработанную с пробелами и хвостом (наименованием, типа 00.0035.ADV-0005 FLIP LED)
        art_predv_2_3 = art_predv_2_1[art_predv_2_1.index('  '):]
        # Избавляемся от серии пробелов, типа '                   00.0035.ADV-0005 FLIP LED'
        art_predv_2_4 = art_predv_2_3.replace('  ','')
        # Зачищаем оставшиеся пробелы и хвост (наименование)
        if art_predv_2_4.index(' ') == 0: #если остался пробел перед артикулом
            art_predv_2_5 = art_predv_2_4[art_predv_2_4.index(' ')+1 :] #перезаписываем без этого первого пробела
            art_predv_2_6 = art_predv_2_5[:art_predv_2_5.index(' ')] #и отбрасываем пробел после артикула и всё что после него
        else: #если перед артикулом пробела нет, отбрасываем последний пробел и всё что после него
            art_predv_2_6 = art_predv_2_4[:art_predv_2_4.index(' ')]
        #Сцепляем первый артикул и две части второго артикула и добавляем в список артикулов
        vendor_code_list.append(art_predv_1 + ' + ' + art_predv_2_2 + art_predv_2_6) # Выцепляем артикул до пробела
    # Если артикул один
    else:
        vendor_code_list.append(art_predv[:art_predv.index(' ')]) # Выцепляем артикул до пробела

    nom_load = nom_load[nom_load.index('Светоотдача')+11: ] #Удаляем инфу по обработанной позиции

number_of_articles = [] #Считаем общее кол-во артикулов
for i in range (len(numbers)):
    if '+' in vendor_code_list[i]: # Если задвоенная позиция, разделяем её и добавляем в общий список артикулов
        vendor_code_0 = vendor_code_list[i]
        vendor_code_1 = vendor_code_0[:vendor_code_0.index('+')-1]
        vendor_code_2 = vendor_code_0[vendor_code_0.index('+')+2 :]
        number_of_articles.append(vendor_code_1)
        number_of_articles.append(vendor_code_2)
    else:
        number_of_articles.append(vendor_code_list[i])

columns = ['Наименование', 'Артикул', 'Кол-во', 'Вход ЭКС, руб', 'Сумма вход, руб', 'МРЦ, руб', 'Сумма МРЦ, руб']
index = [] #Кол-во индексов соответствует кол-ву артикулов
for i in range (len(number_of_articles)):
    index.append(i+1)

data_razdel = []
line_data = []
for i in range (len(vendor_code_list)):
    if '+' in vendor_code_list[i]: #Если есть + в переменной с артикулом(и), то двойная позиция, надо разделить на две строки
        vendor_code_0 = vendor_code_list[i]
        vendor_code_1 = vendor_code_0[:vendor_code_0.index('+')-1]
        vendor_code_2 = vendor_code_0[vendor_code_0.index('+')+2 :]
        try:
            line_data.append(price_varton[price_varton.index == vendor_code_1]['Номенклатура'][0])
            line_data.append(vendor_code_1)
            line_data.append(int(numbers[i]))
            od = float(price_varton[price_varton.index == vendor_code_1]['Вход (за единицу), руб с НДС'][0].replace(' ','').replace(',','.'))
            line_data.append(round(od, 2)) # Вход (за единицу), руб с НДС добавляем в список
            line_data.append(round(float(numbers[i])*od, 2)) # Сумма вход, руб
            mrc = float(price_varton[price_varton.index == vendor_code_1]['МРЦ (за единицу), руб с НДС'][0].replace(' ','').replace(',','.'))
            line_data.append(round(mrc, 2))# МРЦ добавляем в список
            line_data.append(round(float(numbers[i])*mrc, 2)) # Сумма МРЦ, руб
            data_razdel.append(line_data)
            line_data = []
        
        except: #Если нет в прайсе
            line_data1 = ['Нет в прайсе', vendor_code_1, int(numbers[i])]
            data_razdel.append(line_data1)
            line_data = []
    
        try:            
            line_data.append(price_varton[price_varton.index == vendor_code_2]['Номенклатура'][0])
            line_data.append(vendor_code_2)
            line_data.append(int(numbers[i]))
            od = float(price_varton[price_varton.index == vendor_code_2]['Вход (за единицу), руб с НДС'][0].replace(' ','').replace(',','.')) # Вход (за единицу), руб с НДС
            line_data.append(round(od, 2)) # Вход (за единицу), руб с НДС добавляем в список
            line_data.append(round(float(numbers[i])*od, 2)) # Сумма вход, руб
            mrc = float(price_varton[price_varton.index == vendor_code_2]['МРЦ (за единицу), руб с НДС'][0].replace(' ','').replace(',','.'))# МРЦ
            line_data.append(round(mrc, 2))# МРЦ добавляем в список
            line_data.append(round(float(numbers[i])*mrc, 2)) # Сумма МРЦ, руб
            data_razdel.append(line_data)
            line_data = []
        
        except: #Если нет в прайсе
            line_data2 = ['Нет в прайсе', vendor_code_2, int(numbers[i])]
            data_razdel.append(line_data2)
            line_data = []                                                  

    else: #Если нет + в переменной с артикулом(и), то одинарная позиция, разделять не нужно
        try:
            line_data.append(price_varton[price_varton.index == vendor_code_list[i]]['Номенклатура'][0]) # Наименование
            line_data.append(vendor_code_list[i]) # Артикул
            line_data.append(numbers[i]) # Количество
            od = float(price_varton[price_varton.index == vendor_code_list[i]]['Вход (за единицу), руб с НДС'][0].replace(' ','').replace(',','.')) # Вход (за единицу), руб с НДС
            line_data.append(round(od, 2))# Вход (за единицу), руб с НДС добавляем в список
            line_data.append(round(float(numbers[i])*od, 2)) # Сумма вход, руб
            mrc = float(price_varton[price_varton.index == vendor_code_list[i]]['МРЦ (за единицу), руб с НДС'][0].replace(' ','').replace(',','.'))# МРЦ
            line_data.append(round(mrc, 2))# МРЦ добавляем в список
            line_data.append(round(float(numbers[i])*mrc, 2)) # Сумма МРЦ, руб
            data_razdel.append(line_data)
            line_data = []
        except:
            line_data = ['Нет в прайсе', vendor_code_list[i], int(numbers[i])]
            data_razdel.append(line_data)
            line_data = []


df_razdel = pd.DataFrame(data_razdel, columns = columns, index = index) # Создаем ДатаФрейм (в качестве параметров передаем называние столбцов, индексы и сами данные)
print('Кол-во:', round(df_razdel['Кол-во'].sum(), 2))
print('Сумма вход, руб =', round(df_razdel['Сумма вход, руб'].sum(), 2))
print('Сумма МРЦ, руб =', round(df_razdel['Сумма МРЦ, руб'].sum(), 2))
#Добавление строки с суммами по колонкам "вход" и "МРЦ"
df_sums_line = [['', '', '', '', round(df_razdel['Сумма вход, руб'].sum(), 2), '', round(df_razdel['Сумма МРЦ, руб'].sum(), 2)]]
df_sums = pd.DataFrame(df_sums_line, columns = columns)
df_concat = pd.concat([df_razdel, df_sums])

print(df_concat)

writer = pd.ExcelWriter(file_directory_resalts + file_name, engine='xlsxwriter')
df_concat.to_excel(writer, sheet_name='Sheet1') # , index=False
workbook = writer.book #записываем объект 'xlsxwriter' в книгу, для последующих назначений форматов
format1 = workbook.add_format({'num_format': '#,##0.00'})
format2 = workbook.add_format({'num_format': '#,##0.00', 'bold': True})

sheet_0 = writer.sheets['Sheet1']
sheet_0.set_column(0, 0, 5) #индекс
sheet_0.set_column(1, 1, 70) #наименование
sheet_0.set_column(2, 2, 25) #артикул
sheet_0.set_column(3, 3, 7) #кол-во
sheet_0.set_column(4, 4, 14, format1) #цены вход
sheet_0.set_column(5, 5, 15, format2) #сумма вход
sheet_0.set_column(6, 6, 14, format1) #цены МРЦ
sheet_0.set_column(7, 7, 15, format2) #сумма вход

writer.save()