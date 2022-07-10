import openpyxl

wb = openpyxl.load_workbook("transactions.xlsx")
# for printing sheet names
# print(wb.sheetnames)

sheet = wb['Sheet1']

# creating and removing worksheets with index
# wb.create_sheet('Sheet2', 0)
# wb.remove(wb["Sheet2"])
print(wb.sheetnames)

# Accessing an individual cell or a range of cells
# passing coordinate of a cell
cell = sheet['A1'] 
# changing the value of a cell
cell.value = "id"
print(cell.value)


