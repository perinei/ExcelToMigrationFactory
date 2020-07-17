import xlrd
wb_machine_names =  xlrd.open_workbook("Dashboard_06_04.xlsx")
sheet = wb_machine_names.sheet_by_name("Re-host-Dashboard")
# for i in range(5):
#     print (sheet.cell(i+2, 4))

data = [sheet.row_values(i) for i in range(sheet.nrows)]
labels = data[1]    # Don't sort our headers
data = data[2:]     # Data begins on the second row
data.sort(key=lambda x: x[4])

(data[3][3]) = "Pilotao"
print(data[3][3])

for i, value in enumerate(data):
    if value[2] == "R02W02":
        print(f'{value[2]} - {value[4]} - {value[3]}')

