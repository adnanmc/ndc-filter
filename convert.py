import xlrd, json

book = xlrd.open_workbook('product.xlsx')

sheet = book.sheet_by_index(0)

firstRow = sheet.row_values(0)

filteredList = []

for rowIndex in range(1, sheet.nrows):
    currentObject = {}
    
    currentObject[firstRow[1]] = sheet.cell(rowIndex, 1).value
    currentObject[firstRow[2]] = sheet.cell(rowIndex, 2).value
    currentObject[firstRow[3]] = sheet.cell(rowIndex, 3).value
    currentObject[firstRow[4]] = sheet.cell(rowIndex, 4).value
    currentObject[firstRow[5]] = sheet.cell(rowIndex, 5).value
    currentObject[firstRow[6]] = sheet.cell(rowIndex, 6).value
    currentObject[firstRow[7]] = sheet.cell(rowIndex, 7).value
    currentObject[firstRow[13]] = sheet.cell(rowIndex, 13).value
    currentObject[firstRow[14]] = sheet.cell(rowIndex, 14).value
    currentObject[firstRow[15]] = sheet.cell(rowIndex, 15).value
    currentObject["vendor"] = ""
    currentObject["lastOrdered"] = ""
    currentObject["outOfStock"] = ""
    currentObject["orderQuantity"] = ""

    filteredList.append(currentObject)



with open('product.json', 'w') as jsonFile:
    json.dump(filteredList, jsonFile)
