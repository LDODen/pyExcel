import os
import glob
import datetime
import xlrd, xlwt
from xlutils.copy import copy

def findFilesInDir(dir, ext):
    os.chdir(dir)
    result = [i for i in glob.glob('*.{}'.format(ext))]
    return result

print(datetime.date.strftime(datetime.date.today(), "%Y%m%d"))
# res = findFilesInDir('\\\\10.20.0.241\\belatmit\\marshrut_fact\\', 'xls')
res = findFilesInDir('d:\\1\\', 'xls')
print(res)

all = xlrd.open_workbook('d:\\1\\all.xls', on_demand=True, encoding_override="cp1251")
allSheet = all.sheet_by_index(0)
allKolRows = allSheet.nrows
print(allKolRows)

all1 = copy(all)
all1_sheet = all1.get_sheet(0)

all.release_resources()
del all
# имя файла без расширения
# os.path.splittext(res[0])[0]
for i in res:
    xl = xlrd.open_workbook('d:\\1\\' + i, encoding_override="cp1251")
    # x1 = xlrd.open_workbook('d:\\1\\' + str(res[0]), encoding_override="cp1251")
    sheet = xl.sheet_by_index(0)
    sheet_date = sheet.cell(1, 0).value
    print(sheet_date)

    fileIsPresentInAll = False
    for rownum in range(allSheet.nrows):
        if rownum == 0:
            continue
        if (allSheet.row_values(rownum)[0] == sheet_date):
            fileIsPresentInAll = True
            break

    if fileIsPresentInAll == False:
        for rownum in range(sheet.nrows):
            if rownum == 0:
                continue
            print(sheet.cell(rownum, 0).value, rownum)
            all1_sheet.write(allKolRows, 0, sheet.cell(rownum, 0).value)
            all1_sheet.write(allKolRows, 1, sheet.cell(rownum, 1).value)
            all1_sheet.write(allKolRows, 2, sheet.cell(rownum, 2).value)
            all1_sheet.write(allKolRows, 3, sheet.cell(rownum, 3).value)
            all1_sheet.write(allKolRows, 4, sheet.cell(rownum, 4).value)
            all1_sheet.write(allKolRows, 5, sheet.cell(rownum, 5).value)
            all1_sheet.write(allKolRows, 6, sheet.cell(rownum, 6).value)
            allKolRows = allKolRows + 1
        print("file added")
    else:
        print("File present in all")
all1.save('d:\\1\\all.xls')


# vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
# print(vals) #print(sheet.nrows)

# x1 = xlrd.open_workbook('d:\\1\\' + str(i), formatting_info = True)
# print('d:\\1\\' + i)
