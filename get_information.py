import openpyxl
import datetime
from copy import copy

today = datetime.datetime.today()

# today +=  datetime.timedelta(days = 8)

month = today.strftime("%B")
weekday = int(today.strftime("%w"))
weekday += 6
weekday = weekday % 7

day = today.day

print(month, day, weekday)


raw_data = openpyxl.load_workbook('./data/grafik.xlsx')
select_sheet = raw_data['Schedule']
wb = openpyxl.Workbook()
ws = wb.active
workers = {}


# Takes: start cell, end cell, and sheet you want to copy from.
def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    # Loops through selected Rows
    for i in range(startRow, endRow + 1, 1):
        # Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol, endCol + 1, 1):
            rowSelected.append(sheet.cell(row=i, column=j).value)
        # Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)

    return rangeSelected


# Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving, copiedData):
    countRow = 0
    for i in range(startRow, endRow + 1, 1):
        countCol = 0
        for j in range(startCol, endCol + 1, 1):
            sheetReceiving.cell(row=i, column=j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1


def copy_cell(source_cell, row, col, tgt):
    tgt.cell(row=row, column=col).value = source_cell.value
    if source_cell.has_style:
        tgt.cell(row=row, column=col)._style = copy(source_cell._style)


def copyPasteRange(copy_coord, sheet, paste_coord, sheet2):
    startCol, startRow, endCol, endRow = copy_coord[0], copy_coord[1], copy_coord[2], copy_coord[3]
    # Loops through selected Rows
    pasting_col, pasting_row = paste_coord[0], paste_coord[1]
    for i in range(startRow, endRow + 1, 1):
        # Appends the row to a RowSelected list
        for j in range(startCol, endCol + 1, 1):
            cell_to_copy = sheet.cell(row=i, column=j)
            copy_cell(cell_to_copy, pasting_row, pasting_col, sheet2)
            pasting_col += 1
        pasting_row += 1
    return sheet2


for col in range(select_sheet.min_column, select_sheet.max_column + 1):
    if select_sheet.cell(3, col).value == month:
        start_column = col
        break




for row in range(select_sheet.min_row, select_sheet.max_row + 1):
    for col in range(start_column + day - 1 + 7 - weekday, start_column + day - 1 + 7 - weekday + 7):
        if select_sheet.cell(row, col).value == 'x':
            worker_name = select_sheet.cell(row=row, column=1).value
            worker_email = select_sheet.cell(row=row, column=2).value
            workers[worker_name] = worker_email


wb.save('check1.xlsx')
print(workers)
