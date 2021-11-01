from datetime import datetime
import openpyxl

file = openpyxl.load_workbook('result.xlsx')
# print(type(file))
sheets = file.sheetnames
# print(sheets)
# print(file.active.title)
cur_sheet = file['Sheet1']
m = int(input('miqdoriy sonini kiriting: '))
delta_m = int(input('delta miqdorini kiriting: '))
m_miqdoriy = m - delta_m
qoldiq = int(input('qoldiq vagonlar sonini kiriting: '))
values = qoldiq
number_trains = 0
for i in range(781):
    j = 29 + i
    string = "D" + str(j)
    value = cur_sheet[string].value
    values += value
    if values >= m_miqdoriy and values < m:
        remain = 0
        number_trains += 1
        formatted = f'{number_trains}({values})p({remain})q'
        data2 = cur_sheet['F' + str(j)]
        data2.value = formatted
        file.save('result.xlsx')
        values = remain
    if values > m:
        remain = values - m
        number_trains += 1
        formatted = f'{number_trains}p({m})({remain})q'
        data2 = cur_sheet['F' + str(j)]
        data2.value = formatted
        file.save('result.xlsx')
        values = remain
    print(number_trains, j, values)

print('Ikkinchi algoritm boshlandi')

for i in range(781):
    j = 29 + i
    string = "D" + str(j)
    value = cur_sheet[string].value
    values += value
    if values >= m:
        remain = values - m
        number_trains += 1
        formatted = f'{number_trains}p({remain})q'
        data2 = cur_sheet['E' + str(j)]
        data2.value = formatted
        file.save('result.xlsx')
        values = remain
    print(number_trains, j, values)


