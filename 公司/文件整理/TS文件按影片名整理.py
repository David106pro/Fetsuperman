# -*- coding: utf-8 -*-
import os
import shutil
import sys

def organize_ts_files(target_directory):
    """
    该函数用于整理指定目录下的.ts文件。
    它会根据文件名中第一个下划线前的内容（影片名称）创建文件夹，
    并将具有相同影片名称的.ts文件移动到对应的文件夹中。
    
    :param target_directory: 包含.ts文件的目录路径
    """
    # 步骤1：获取所有.ts文件
    try:
        # 获取目录下所有文件和文件夹的名称
        all_entries = os.listdir(target_directory)
        # 筛选出.ts文件，并确保它们是文件而不是目录
        ts_files = [f for f in all_entries if f.endswith('.ts') and os.path.isfile(os.path.join(target_directory, f))]
    except FileNotFoundError:
        print(f"错误：找不到目录 '{os.path.abspath(target_directory)}'。")
        return
    except Exception as e:
        print(f"读取目录 '{os.path.abspath(target_directory)}' 时出错: {e}")
        return

    if not ts_files:
        print(f"在目录 '{os.path.abspath(target_directory)}' 中没有找到.ts文件。")
        return

    # 步骤2：根据影片名称对文件进行分组
    movies = {} # 创建一个空字典来存储影片和对应的文件列表
    for filename in ts_files:
        # 检查文件名中是否包含下划线
        if '_' in filename:
            # 使用第一个下划线分割文件名，获取影片名称
            movie_name = filename.split('_')[0]
            # 如果字典中还没有这个影片名称的键，则创建一个
            if movie_name not in movies:
                movies[movie_name] = []
            # 将当前文件名添加到对应影片的列表中
            movies[movie_name].append(filename)

    print(f"在 '{os.path.abspath(target_directory)}' 中发现了 {len(ts_files)} 个.ts文件, 将整理为 {len(movies)} 个影片文件夹。")

    # 步骤3：创建文件夹并移动文件
    for movie_name, files in movies.items():
        # 新文件夹的路径
        new_folder_path = os.path.join(target_directory, movie_name)
        try:
            # 创建影片文件夹，如果文件夹已存在则不报错
            os.makedirs(new_folder_path, exist_ok=True)
            print(f"已创建或确认文件夹: {movie_name}")

            # 将文件移动到新创建的文件夹中
            for file_to_move in files:
                source_path = os.path.join(target_directory, file_to_move)
                destination_path = os.path.join(new_folder_path, file_to_move)
                # 移动文件
                if os.path.exists(source_path):
                    shutil.move(source_path, destination_path)
                    print(f"  已移动: {file_to_move} -> {movie_name}{os.sep}")
                else:
                    print(f"  警告: 文件 {file_to_move} 在移动前已经不存在于源目录。")
        except OSError as e:
            print(f"为影片 '{movie_name}' 创建目录或移动文件时出错: {e}")
        except Exception as e:
            print(f"处理影片 '{movie_name}' 时发生未知错误: {e}")

    print("\n文件整理完成。")

if __name__ == "__main__":
    # 检查是否通过命令行传递了目录参数
    if len(sys.argv) > 1:
        # 使用命令行提供的第一个参数作为目标目录
        directory = sys.argv[1]
    else:
        # 如果没有提供参数，则使用指定的电影目录
        directory = r'E:\8M\电影'
        print(f"未指定目录，将使用目标目录: {directory}")

    # 调用主函数，并传入目标目录
    organize_ts_files(directory)