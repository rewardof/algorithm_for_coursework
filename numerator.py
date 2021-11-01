# s = 'brian-23@hackerrank.com'
# name = s[:s.find('@')]
# website = s[s.find('@')+1:s.find('.')]
# extension = s[s.find('.')+1:]
# print(name)
# print(website)
# print(extension)
import openpyxl

file = openpyxl.load_workbook('main_doc.xlsx')
# print(type(file))
sheets = file.sheetnames
# print(sheets)
# print(file.active.title)
cur_sheet = file['Sheet1']
print(cur_sheet['G190'].value)
if cur_sheet['G190'].value is None:
    print('yes')