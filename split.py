import openpyxl

# 打开Excel文件
workbook = openpyxl.load_workbook(r'D:\Personal\Desktop\DDI_PRoject\lung_Cominfo.xlsx')
sheet=workbook['DCDB']

# 遍历每一行数据（从第二行开始，假设第一行是标题行）
for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=False), start=2):
    # 获取第2列的数据
    second_column_data = row[1].value

    if second_column_data:
        # 以"; "为分隔符拆分数据
        split_data = second_column_data.split("; ")

        # 将拆分后的数据分别存入第2列和第3列
        row[1].value = split_data[0]
        if len(split_data) > 1:
            row[2].value = split_data[1]

# 保存修改后的Excel文件
workbook.save(r'D:\Personal\Desktop\DDI_PRoject\lung_Cominfo.xlsx')
