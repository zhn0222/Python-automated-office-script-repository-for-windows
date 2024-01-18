import os
import pandas as pd
import win32com.client as win32

def convert_to_xlsx(input_path, output_path):
    print(f"正在处理文件: {input_path}")
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(input_path)
        wb.SaveAs(output_path, FileFormat=51)
        wb.Close()
        excel.Quit()
        print(f"转换完成: {output_path}")
    except Exception as e:
        print(f"转换文件时发生错误: {str(e)}")

def convert_all_to_xlsx(root_folder, output_root):
    all_items = os.listdir(root_folder)

    xls_files = [item for item in all_items if item.endswith('.xls')]
    for xls_file in xls_files:
        input_path = os.path.join(root_folder, xls_file)
        output_file = os.path.splitext(xls_file)[0] + '.xlsx'
        output_path = os.path.join(output_root, output_file)

        output_folder = os.path.join(output_root, os.path.basename(root_folder))
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        convert_to_xlsx(input_path, output_path)

    subfolders = [item for item in all_items if os.path.isdir(os.path.join(root_folder, item))]
    for subfolder in subfolders:
        folder_path = os.path.join(root_folder, subfolder)
        output_subfolder = os.path.join(output_root, subfolder)

        if not os.path.exists(output_subfolder):
            os.makedirs(output_subfolder)

        subfolder_xls_files = [f for f in os.listdir(folder_path) if f.endswith('.xls')]

        for xls_file in subfolder_xls_files:
            input_path = os.path.join(folder_path, xls_file)
            output_file = os.path.splitext(xls_file)[0] + '.xlsx'
            output_path = os.path.join(output_subfolder, output_file)
            convert_to_xlsx(input_path, output_path)

if __name__ == "__main__":
    # 指定根文件夹和输出根文件夹的路径
    root_folder = r''
    output_root = r''

    convert_all_to_xlsx(root_folder, output_root)

    print("批量转换完成！")