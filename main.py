import xlrd
import xlwt
book = xlrd.open_workbook('ATD-EFI8ST-EW-V1.4.xls')
sh = book.sheet_by_index(0)

sum_text = ""
dataJson = []
tempDesignator = []
number = 0
for i in range(sh.nrows-1):
    if len(dataJson) == 0:
        dataJson.append({
            "part_type":sh.cell_value(rowx=i+1, colx=0),
            "designator":sh.cell_value(rowx=i+1, colx=1),
            "foot_print":sh.cell_value(rowx=i+1, colx=2),
            "description":sh.cell_value(rowx=i+1, colx=3)
        })
        number += 1
    else:
        if dataJson[number-1]['part_type'] == sh.cell_value(rowx=i+1, colx=0) and dataJson[number-1]['foot_print'] == sh.cell_value(rowx=i+1, colx=2):
            dataJson[number-1]['designator'] = dataJson[number-1]['designator']+','+sh.cell_value(rowx=i+1, colx=1)
        else:
            dataJson.append({
                "part_type":sh.cell_value(rowx=i+1, colx=0),
                "designator":sh.cell_value(rowx=i+1, colx=1),
                "foot_print":sh.cell_value(rowx=i+1, colx=2),
                "description":sh.cell_value(rowx=i+1, colx=3)
            })
            number += 1
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('output')
number = 0
for i in dataJson:
    sheet.write(number, 0, i['part_type'])
    sheet.write(number, 1, i['designator'])
    sheet.write(number, 2, i['foot_print'])
    sheet.write(number, 3, i['description'])
    number+=1

workbook.save('output.xls')