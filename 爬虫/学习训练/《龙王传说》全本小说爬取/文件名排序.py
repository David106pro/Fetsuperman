import os

def rename_files(folder_path):
    files = os.listdir(folder_path)
    chapter_files = [f for f in files if f.endswith('.txt') and '第' in f]
    chapter_files.sort(key=lambda x: try_to_convert(x))

    for index, file in enumerate(chapter_files, 1):
        new_name = f"第{index}章.txt"
        os.rename(os.path.join(folder_path, file), os.path.join(folder_path, new_name))

def try_to_convert(x):
    try:
        return int(x.split('第')[1].split('章')[0])
    except ValueError:
        # 如果无法转换为整数，返回一个很大的数，使其排在后面
        return 999999