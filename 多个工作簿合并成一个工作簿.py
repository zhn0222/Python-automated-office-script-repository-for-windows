import os
import win32com.client as win32

def merge_excel_files(input_folder, output_file):
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    merged_workbook = excel.Workbooks.Add()

    excel_files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx")]

    for excel_file in excel_files:
        current_workbook = excel.Workbooks.Open(os.path.join(input_folder, excel_file))

        for sheet in current_workbook.Sheets:
            sheet.Copy(Before=merged_workbook.Sheets(1))

        current_workbook.Close()

    for sheet in merged_workbook.Sheets:
        if sheet.UsedRange.Address == "$A$1":
            sheet.Delete()

    merged_workbook.SaveAs(output_file)
    merged_workbook.Close()

    excel.Application.Quit()

# 输入文件夹路径和输出文件路径
input_folder = r""
output_excel_file = r""

merge_excel_files(input_folder, output_excel_file)