import openpyxl
workbook = openpyxl.load_workbook('D:\Personal\Desktop\DDI_PRoject\DB_ID与特征\DB_Name\DB_Name.xlsx')
sheet = workbook.active
for row in sheet.iter_rows(min_row=2, values_only=False):
    data_str = row[0].value
    split_data = data_str.split(",", 4)
    for i, value in enumerate(split_data[:4], start=2):
        row[i].value = value
    row[6].value = split_data[-1]
workbook.save('D:\Personal\Desktop\DDI_PRoject\DB_ID与特征\DB_Name\DB_Name.xlsx')
