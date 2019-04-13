import openpyxl as xl
from openpyxl.chart import BarChart, Reference

my_wb = xl.load_workbook('csldata.xlsx')
first_sheet = my_wb['Sheet1']
cell = first_sheet.cell(1,1)

#print(cell.value) to get the item in a cell at a particular coordinate
#print(first_sheet.max_row) to get the total number of rows in a sheet



for row in range(2, first_sheet.max_row+1): #i started from 2 to ignore headings and incremented 1 because range function is size -1
    cell = first_sheet.cell(row, 3)
    updated_marks = cell.value + 5
    updated_marks_cell = first_sheet.cell(row, 4)
    updated_marks_cell.value = updated_marks


values = Reference(first_sheet, min_row=2, max_row=first_sheet.max_row, min_col=4, max_col=4)

chart = BarChart()
chart.add_data(values)
first_sheet.add_chart(chart, 'e2')


my_wb.save('csldata_withChart.xlsx')
