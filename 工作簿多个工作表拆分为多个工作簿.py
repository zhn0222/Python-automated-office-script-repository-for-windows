import win32com.client as win32

def split_excel_by_sheet(input_file, output_folder):
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    workbook = excel.Workbooks.Open(input_file)

    sheet_names = [sheet.Name for sheet in workbook.Sheets]

    for sheet_name in sheet_names:
        sheet = workbook.Sheets(sheet_name)

        new_workbook = excel.Workbooks.Add()
        sheet.Copy(Before=new_workbook.Sheets(1))

        output_file = f"{output_folder}/{sheet_name}.xlsx"
        new_workbook.SaveAs(output_file)
        new_workbook.Close()

        print(f"工作表 '{sheet_name}' 已拆分为 '{output_file}'")

    workbook.Close()

    excel.Application.Quit()

# 输出文件夹路径
output_folder = r""

# 输入文件夹路径
input_excel_file = r""
split_excel_by_sheet(input_excel_file, output_folder)