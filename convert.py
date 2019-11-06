import xlrd, json

book = xlrd.open_workbook('product.xlsx')

sheet = book.sheet_by_index(0)

firstRow = sheet.row_values(0)

for columnId in range(len(firstRow)):
    print(f"{firstRow[columnId]} - {columnId}")

dict_list = []

for rowIndex in range(1, sheet.nrows):
    dict_object = {}
    # for colIndex in range(len(firstRow)):
    #     dict_object[firstRow[colIndex]] = sheet.cell(rowIndex, colIndex).value
    dict_object[firstRow[0]] = sheet.cell(rowIndex, 0).value
    dict_object[firstRow[0]] = sheet.cell(rowIndex, 0).value
    dict_object[firstRow[0]] = sheet.cell(rowIndex, 0).value
    dict_object[firstRow[0]] = sheet.cell(rowIndex, 0).value

    dict_list.append(dict_object)


# print(dict_list)

with open('product.json', 'w') as jsonFile:
    json.dump(dict_list, jsonFile)
