import os
import openpyxl


def xlsx_to_csv_batch(folder_path):
    """
    批量将指定文件夹下的所有xlsx文件转换为csv文件
    """
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xlsx'):
                xlsx_file_path = os.path.join(root, file)
                csv_file_path = os.path.splitext(xlsx_file_path)[0] + '.csv'
                try:
                    workbook = openpyxl.load_workbook(xlsx_file_path)
                    try:
                        sheet = workbook.active
                        with open(csv_file_path, 'w', newline='') as csv_file:
                            for row in sheet.rows:
                                row_data = []
                                for cell in row:
                                    row_data.append(str(cell.value))
                                csv_file.write(','.join(row_data) + '\n')
                    except SomeInnerException as inner_e:  # 捕获内层可能出现的特定异常
                        print(f"内层操作出现错误: {inner_e}")
                except Exception as e:  # 捕获外层整体操作可能出现的异常
                    print(f"转换 {xlsx_file_path} 时出现错误: {e}")



if __name__ == "__main__":
    folder_path = r"D:\Personal\Desktop\DDI_PROject\IEEE JBHI 2024Winter\ALL DATA\Similarity\ATC"  # 替换为实际的目标文件夹路径
    xlsx_to_csv_batch(folder_path)
    
