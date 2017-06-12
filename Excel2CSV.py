import xlrd
import csv
from os import sys

def csv_from_excel(excel_file):
    workbook = xlrd.open_workbook(excel_file)
    all_worksheets = workbook.sheet_names()

    for worksheet_name in all_worksheets:
        worksheet = workbook.sheet_by_name(worksheet_name)
        your_csv_file = open(''.join([worksheet_name,'.csv']), 'w', newline = '')
        wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

        for rownum in range(worksheet.nrows):
            value = str ( worksheet.row (rownum)[1].value )
            wr.writerow(x for x in worksheet.row_values(rownum))#list(x.encode('utf-8') if type(x) == type(u'') else x for x in worksheet.row_values(rownum)))
        your_csv_file.close()

if __name__ == "__main__":
  csv_from_excel('C:/Users/smitha/Downloads/PJM_V4.xlsx')
