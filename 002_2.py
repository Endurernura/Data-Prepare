import openpyxl

def process_sheet_data():
    input_file_path = r'D:\Personal\Desktop\DDI_PROject\OUTPUT\FINAL_Prepare\SMILES\rtcb1.xlsx'
    output_file_path = r'D:\Personal\Desktop\DDI_PROject\OUTPUT\FINAL_Prepare\SMILES\rtcb2.xlsx'
    workbook = openpyxl.load_workbook(input_file_path)
    sheet = workbook['cdcdb']

    # 用于记录第一列数据对应的最小行号
    first_column_min_rows = {}
    # 用于存储需要删除的行号
    rows_to_delete = []

    # 遍历每一行数据（从第二行开始）
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) >= 1:
            first_column_value = row[0]
            if first_column_value in first_column_min_rows:
                if first_column_min_rows[first_column_value] > row_index:
                    # 如果当前行号更小，更新最小行号，并把之前记录的最小行号对应的行加入待删除列表
                    rows_to_delete.append(first_column_min_rows[first_column_value])
                    first_column_min_rows[first_column_value] = row_index
                else:
                    # 如果当前行号更大，将当前行加入待删除列表
                    rows_to_delete.append(row_index)
            else:
                first_column_min_rows[first_column_value] = row_index

    # 按照行号从大到小的顺序删除行，避免索引混乱
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    workbook.save(output_file_path)


if __name__ == '__main__':
    process_sheet_data()