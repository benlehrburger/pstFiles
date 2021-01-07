# Find emails in an Excel file

from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

set1 = {'email@email.com'}
set2 = {'email2@email.com'}

master = set1 | set2

row = 0
for i in master:
    sheet1.write(row, 0, i)
    row += 1

wb.save('Emails.xls')
