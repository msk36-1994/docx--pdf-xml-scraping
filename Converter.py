# -*- coding: utf-8 -*-

import os, os.path
from datetime import datetime
import zipfile, fnmatch
import xml.etree.ElementTree as ET
from docx2pdf import convert

count_files_in_dir = len(fnmatch.filter(os.listdir('./docx'), 'result*'))
time_of_run = count_files_in_dir * 3
print(f'Примерное время выполнения скрипта составляет: {time_of_run} минут(ы)')

uk_name = input('Введите имя УК: ')
print('Начато выполнение скрипта')
start_time = datetime.now()

dr = r'./docx'
test = os.listdir(dr)
list1 = []
list2 = []
k = 0
i = 0
extract_dir = r'.\Остальное\docx_tmp\word'
file_pattern = 'word/footer1.xml'

for item in test:
    print('----------------------------------------------------------')
    start_time1 = datetime.now()

    try:
        print(f'Открываем {item}')
        with zipfile.ZipFile(
                f'\\\\guo.local\\DFSFILES\\MSK\HOME\\dobychin.ia\\Desktop\\Задания\\docx-pdf\\docx\\{item}', 'r') as zf:
            for file in zf.infolist():
                if fnmatch.fnmatch(file.filename, file_pattern):
                    zf.extract(file.filename, extract_dir)
            zf.close()
    except:
        print(f'{item} не подлежит распаковке')

    tree = ET.parse(
        r'\\guo.local\DFSFILES\MSK\HOME\dobychin.ia\Desktop\Задания\docx-pdf\Остальное\docx_tmp\word\word\footer1.xml')
    root = tree.getroot()
    lst = tree.findall('ftr/p/r/t')
    counts = tree.findall('.//')
    list1.clear()

    for each in counts:
        try:
            list1.append(each.text)
        except:
            pass

    xml_data = list(filter(lambda e: e != None, list1))

    index = xml_data[0].split(',')[0]
    UK = xml_data[0]

    try:
        UK1 = index[i]
    except:
        pass

    list2.clear()
    list2.append(UK)

    try:
        convert(f'.//docx//{item}', f'.//pdf//{uk_name} {list2[0]}.pdf')
        print(f'Время обработки {item}: {datetime.now() - start_time1}')
        print('----------------------------------------------------------')
    except:
        break

    i = i + 1
    continue

print(f'Время выполнения скрипта: {datetime.now() - start_time}')
