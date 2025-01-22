import openpyxl

def remove_blank_rows():
    workbook = openpyxl.load_workbook(r'D:\Personal\Desktop\DDI_PROject\OUTPUT\FINAL_Prepare\SMILES\rtcb.xlsx')
    sheet = workbook['drugmap']

    rows_to_delete = []
    for row_index, row in enumerate(sheet.iter_rows(min_row = 2, values_only=True), start = 2):

        dbid1 = row[0]

        if dbid1 is None:
            rows_to_delete.append(row_index)

    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    workbook.save(r'D:\Personal\Desktop\DDI_PROject\OUTPUT\FINAL_Prepare\SMILES\rtcb.xlsx')

if __name__ == "__main__":
    remove_blank_rows()
