import os
import xlrd
import xlwt
#mystr=os.popen("pdf2htmlEX.exe 1.pdf")





#写入
def Write2Excel(names):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('My Worksheet')
    for row in range(1, len(names)+1):
        worksheet.write(row, 0, names[row-1])
    workbook.save('out.xls')



rootdif = 'pdfs2html'
lists = os.listdir(rootdif)
for list in lists:
    cmd = 'pdf2htmlEX.exe pdfs2html/' + list
    print(cmd)
    os.system(cmd)
