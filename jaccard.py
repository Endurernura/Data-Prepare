import openpyxl


def calculate_jaccard_similarity():
    workbook = openpyxl.load_workbook(r'D:\Personal\Desktop\ALL DATA\Similarity\ATC\ATC_Breast.xlsx')
    sheet = workbook['atc']

    # 获取第一行数据（除第一列，假设第一列是表头）
    first_row_data = [cell.value for cell in sheet['B1:GU1'][0]]

    # 遍历第一列数据（从第二行开始，假设第一行是表头）
    for row_index, cell in enumerate(sheet['A'][1:], start=2):
        row_data = cell.value
        similarity_results = []

        for col_data in first_row_data:
            if row_data and col_data:
                set1 = set(row_data[:4])
                set2 = set(col_data[:4])

                intersection = len(set1.intersection(set2))
                union = len(set1.union(set2))

                similarity = intersection / union if union!= 0 else 0
            else:
                similarity = 0

            similarity_results.append(similarity)

        # 将相似性结果输出到对应行的第二列到最后一列
        for col_index, similarity_value in enumerate(similarity_results, start=2):
            sheet.cell(row=row_index, column=col_index).value = similarity_value

    workbook.save(r'D:\Personal\Desktop\ALL DATA\Similarity\ATC\ATC_Breast.xlsx')


if __name__ == '__main__':
    calculate_jaccard_similarity()
