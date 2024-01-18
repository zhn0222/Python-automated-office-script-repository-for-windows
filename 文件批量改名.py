import os

def batch_rename_recursive(directory_path, old_part, new_part):
    for root, dirs, files in os.walk(directory_path):
        for file_name in files:
            if old_part in file_name:
                new_file_name = file_name.replace(old_part, new_part)

                old_path = os.path.join(root, file_name)
                new_path = os.path.join(root, new_file_name)

                os.rename(old_path, new_path)
                print(f'Renamed: {file_name} to {new_file_name}')

# 替换以下路径、旧部分和新部分为你自己的值
directory_path = r''  # 指定主文件夹路径
old_part = ''  # 要替换的旧部分
new_part = ''  # 替换为的新部分

batch_rename_recursive(directory_path, old_part, new_part)