from datetime import datetime

import openpyxl

file = openpyxl.load_workbook('main_doc.xlsx')
sheets = file.sheetnames
cur_sheet = file['Sheet1']
l = []
l1 = []
l3 = []
l7 = []
l13 = []
time3 = '03:00:00'
time7 = '07:00:00'
time13 = '13:00:00'
time17 = '17:00:00'
time23 = '23:00:00'
time01 = '0:01:00'

time3 = datetime.strptime(time3, "%H:%M:%S").time()
time7 = datetime.strptime(time7, "%H:%M:%S").time()
time13 = datetime.strptime(time13, "%H:%M:%S").time()
time17 = datetime.strptime(time17, "%H:%M:%S").time()
time23 = datetime.strptime(time23, "%H:%M:%S").time()
time01 = datetime.strptime(time01, "%H:%M:%S").time()
m = int(input('miqdoriy sonini kiriting: '))
delta_m = int(input('delta miqdorini kiriting: '))
m_miqdoriy = m - delta_m
ikki_m = 2 * m
ikki_m_delta = 2 * m - delta_m
uch_m = 3 * m
uch_m_delta = 3 * m - delta_m
qoldiq = int(input('qoldiq vagonlar sonini kiriting: '))
values = qoldiq
number_trains = 0
j = 29
for i in range(781):
    string1 = 'C' + str(j)
    string2 = 'D' + str(j)
    string3 = 'G' + str(j)
    time = cur_sheet[string1].value
    time1 = time
    time = str(time)
    time = datetime.strptime(time, "%H:%M:%S").time()
    if time is None:
        break
    value = cur_sheet[string2].value
    total = (time1, value, i)
    values += value
    print(values)
    l.append(total)
    color = cur_sheet[string3].fill.start_color.index
    if color == 'FF0070C0':
        if values < m:
            formatted = f'{values} - otmena'
            data3 = cur_sheet[string3]
            data3.value = formatted
        if values >= m and values < ikki_m:
            row = cur_sheet.insert_rows(j + 1, 1)
            file.save('main_doc.xlsx')
            remain = values - m
            number_trains += 1
            formatted = f'{number_trains}p({remain})q'
            data2 = cur_sheet['C' + str(j + 1)]
            if time > time01 and time < time3:
                data2.value = '3:00'
            if time > time3 and time < time7:
                data2.value = '7:00'
            if time > time7 and time < time13:
                data2.value = '13:00'
            if time > time13 and time < time17:
                data2.value = '17:00'
            if time > time17 and time < time23:
                data2.value = '23:00'
            data4 = cur_sheet['A' + str(j + 1)]
            data4.value = cur_sheet['A' + str(j)].value
            data3 = cur_sheet['G' + str(j + 1)]
            data3.value = formatted
            values = remain
            print(values)
            j = j + 1
        if values >= ikki_m and values < uch_m:
            row = cur_sheet.insert_rows(j + 1, 1)
            file.save('main_doc.xlsx')
            remain = values - ikki_m
            number_trains += 2
            formatted = f'{number_trains}p({remain})q'
            data2 = cur_sheet['C' + str(j + 1)]
            if time > time01 and time < time3:
                data2.value = '3:00'
            if time > time3 and time < time7:
                data2.value = '7:00'
            if time > time7 and time < time13:
                data2.value = '13:00'
            if time > time13 and time < time17:
                data2.value = '17:00'
            if time > time17 and time < time23:
                data2.value = '23:00'
            data4 = cur_sheet['A' + str(j + 1)]
            data4.value = cur_sheet['A' + str(j)].value
            data3 = cur_sheet['G' + str(j + 1)]
            data3.value = formatted
            values = remain
            print(values)
            j = j + 1
        if values >= uch_m:
            remain = values - uch_m
            number_trains += 3
            formatted = f'{number_trains}p({remain})q'
            data3 = cur_sheet['G' + str(j + 1)]
            data3.value = formatted
            values = remain
            print(values)
            print('u 180 dan oshdi')
    j = j + 1
number_row = 0
for i in range(2, 10000):
    if cur_sheet['A' + str(i)].value is None:
        number_row = i
        break

num_train = 0
values1 = qoldiq
values2 = qoldiq
j = 29
print('2 - algoritm boshlandi')
for i in range(29, number_row):
    string2 = 'D' + str(j)
    string3 = 'H' + str(j)
    value = cur_sheet[string2].value
    if not value:
        value = 0
    values1 += value
    values2 += value
    print(values2, values1, value, j, num_train + 1)
    color = cur_sheet[string3].fill.start_color.index
    if color == 'FF0070C0':
        if values1 < m_miqdoriy:
            formatted = f'{values1} - otmena'
            data3 = cur_sheet[string3]
            data3.value = formatted
            file.save('main_doc.xlsx')
            if cur_sheet['G' + str(j)].value is None:
                j = j + 1

        if values1 >= m_miqdoriy and values1 <= m:
            if cur_sheet['D' + str(j + 1)].value:
                print('this is not empty')
                row = cur_sheet.insert_rows(j + 1, 1)
                file.save('main_doc.xlsx')
                remain = 0
                num_train += 1
                formatted = f'{num_train}p({values1})({remain})q'
                data3 = cur_sheet['H' + str(j + 1)]
                data3.value = formatted
                file.save('main_doc.xlsx')
                values1 = remain
                j = j + 1
            else:
                remain = 0
                num_train += 1
                formatted = f'{num_train}p({values1})({remain})q'
                data3 = cur_sheet['H' + str(j + 1)]
                data3.value = formatted
                file.save('main_doc.xlsx')
                values1 = remain
                j = j + 1

        if values1 > m and values1 < ikki_m_delta:
            if cur_sheet['D' + str(j + 1)].value:
                print('this is not empty')
                row = cur_sheet.insert_rows(j + 1, 1)
                file.save('main_doc.xlsx')
                remain = values1 - m
                num_train += 1
                formatted = f'{num_train}p({m})({remain})q'
                data3 = cur_sheet['H' + str(j + 1)]
                data3.value = formatted
                file.save('main_doc.xlsx')
                values1 = remain
                j = j + 1
            else:
                remain = values1 - m
                num_train += 1
                formatted = f'{num_train}p({m})({remain})q'
                data3 = cur_sheet['H' + str(j + 1)]
                data3.value = formatted
                file.save('main_doc.xlsx')
                values1 = remain
                j = j + 1

        if values1 >= ikki_m_delta:
            do_value = values1 - m
            if do_value >= m_miqdoriy and do_value <= m:
                if cur_sheet['D' + str(j + 1)].value:
                    print('this is not empty')
                    row = cur_sheet.insert_rows(j + 1, 1)
                    file.save('main_doc.xlsx')
                    remain = 0
                    num_train += 2
                    formatted = f'{num_train}p({values1 - m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
                else:
                    remain = 0
                    num_train += 2
                    formatted = f'{num_train}p({values1 - m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
            if do_value > m and do_value < ikki_m_delta:
                if cur_sheet['D' + str(j + 1)].value:
                    print('this is not empty')
                    row = cur_sheet.insert_rows(j + 1, 1)
                    file.save('main_doc.xlsx')
                    remain = values1 - ikki_m
                    num_train += 2
                    formatted = f'{num_train}p({m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
                else:
                    remain = values1 - ikki_m
                    num_train += 2
                    formatted = f'{num_train}p({m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
        if values1 >= uch_m_delta:
            do_value = values1 - ikki_m
            if do_value >= m_miqdoriy and do_value <= m:
                if cur_sheet['D' + str(j + 1)].value:
                    print('this is not empty')
                    row = cur_sheet.insert_rows(j + 1, 1)
                    file.save('main_doc.xlsx')
                    remain = 0
                    num_train += 3
                    formatted = f'{num_train}p({values1 - ikki_m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
                else:
                    remain = 0
                    num_train += 3
                    formatted = f'{num_train}p({values1 - ikki_m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
            if do_value > m and do_value < ikki_m_delta:
                if cur_sheet['D' + str(j + 1)].value:
                    print('this is not empty')
                    row = cur_sheet.insert_rows(j + 1, 1)
                    file.save('main_doc.xlsx')
                    remain = values1 - uch_m
                    num_train += 3
                    formatted = f'{num_train}p({m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
                else:
                    remain = values1 - uch_m
                    num_train += 3
                    formatted = f'{num_train}p({m})({remain})q'
                    data3 = cur_sheet['H' + str(j + 1)]
                    data3.value = formatted
                    file.save('main_doc.xlsx')
                    values1 = remain
                    j = j + 1
    j = j + 1
    if j == number_row:
        break
