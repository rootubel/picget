import os
import zipfile
from PIL import Image
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def extract_first_image_from_zip(zip_path):
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            # 找出图片文件（jpg/jpeg/png等）
            imgs = [f for f in z.namelist() if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif'))]
            if imgs:
                # 把图片解压到内存，保存临时文件再打开（简化处理）
                temp_img_name = imgs[0]
                with z.open(temp_img_name) as img_file:
                    img = Image.open(img_file)
                    return img
    except Exception as e:
        print(f"解压zip失败：{zip_path}，错误：{e}")
    return None

def extract_first_image_from_folder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif')):
                img_path = os.path.join(root, file)
                try:
                    img = Image.open(img_path)
                    return img
                except:
                    continue
    return None

def main():
    root = tk.Tk()
    root.withdraw()

    # 选文件夹，里面有zip或者普通文件夹
    input_dir = filedialog.askdirectory(title="选择包含压缩包或文件夹的目录")
    if not input_dir:
        print("未选择输入目录，程序退出。")
        return

    # 选款色表（xls/xlsx/csv）
    excel_path = filedialog.askopenfilename(title="选择款色表(Excel或CSV)", filetypes=[("Excel files","*.xlsx *.xls"),("CSV files","*.csv")])
    if not excel_path:
        print("未选择款色表，程序退出。")
        return

    # 读取表格
    try:
        if excel_path.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(excel_path)
        else:
            df = pd.read_csv(excel_path, encoding='utf-8')
    except Exception as e:
        print(f"读取表格失败：{e}")
        return

    # 表格列名假设（根据你的表格调整）
    # 款号 - '款号'
    # 色号 - '色号'
    # 颜色描述 - '颜色'

    # 建立款号-颜色描述映射，方便查找色号
    color_map = {}
    for _, row in df.iterrows():
        key = (str(row['款号']).strip(), str(row['颜色']).strip())
        color_map[key] = str(row['色号']).strip()

    # 输出目录
    output_dir_style = os.path.join(input_dir, "款号图")
    output_dir_color = os.path.join(input_dir, "款色图")
    os.makedirs(output_dir_style, exist_ok=True)
    os.makedirs(output_dir_color, exist_ok=True)

    processed_styles = set()  # 存已生成款号图，避免重复

    # 遍历目录，支持zip和普通文件夹
    for item in os.listdir(input_dir):
        full_path = os.path.join(input_dir, item)
        if item.lower().endswith('.zip'):
            # zip文件，款号通常是zip名开头部分
            style_code = item.split()[0]
            if style_code in processed_styles:
                print(f"款号图已存在，跳过款号图: {style_code}")
            else:
                img = extract_first_image_from_zip(full_path)
                if img:
                    save_path = os.path.join(output_dir_style, f"{style_code}.jpg")
                    if not os.path.exists(save_path):
                        img.save(save_path)
                        print(f"保存款号图: {style_code}.jpg")
                        processed_styles.add(style_code)
                    else:
                        print(f"款号图已存在，跳过保存: {style_code}.jpg")
                else:
                    print(f"未找到zip中图片: {item}")

            # 进一步寻找款色文件夹，示例假设解压后有同名文件夹，且颜色文件夹里有图片
            folder_inside = os.path.join(input_dir, style_code)
            if os.path.exists(folder_inside) and os.path.isdir(folder_inside):
                for color_folder in os.listdir(folder_inside):
                    color_path = os.path.join(folder_inside, color_folder)
                    if os.path.isdir(color_path):
                        # color_folder名一般含颜色描述，匹配表格
                        color_desc = color_folder.strip()
                        key = (style_code, color_desc)
                        if key in color_map:
                            color_code = color_map[key]
                            save_name = f"{style_code}_{color_code}_IMAGEURL.jpg"
                            save_path = os.path.join(output_dir_color, save_name)
                            if os.path.exists(save_path):
                                print(f"款色图已存在，跳过: {save_name}")
                            else:
                                img_color = extract_first_image_from_folder(color_path)
                                if img_color:
                                    img_color.save(save_path)
                                    print(f"保存款色图: {save_name}")
                                else:
                                    print(f"款色文件夹无图片: {color_path}")
                        else:
                            print(f"⚠️ 未匹配款色：{style_code} / {color_desc}")
            else:
                print(f"未找到对应文件夹: {folder_inside}")

        elif os.path.isdir(full_path):
            # 普通文件夹，款号通常是文件夹名开头部分
            style_code = item.split()[0]
            if style_code in processed_styles:
                print(f"款号图已存在，跳过款号图: {style_code}")
            else:
                img = extract_first_image_from_folder(full_path)
                if img:
                    save_path = os.path.join(output_dir_style, f"{style_code}.jpg")
                    if not os.path.exists(save_path):
                        img.save(save_path)
                        print(f"保存款号图: {style_code}.jpg")
                        processed_styles.add(style_code)
                    else:
                        print(f"款号图已存在，跳过保存: {style_code}.jpg")
                else:
                    print(f"未找到文件夹中图片: {item}")

            # 找子文件夹（颜色文件夹）
            for color_folder in os.listdir(full_path):
                color_path = os.path.join(full_path, color_folder)
                if os.path.isdir(color_path):
                    color_desc = color_folder.strip()
                    key = (style_code, color_desc)
                    if key in color_map:
                        color_code = color_map[key]
                        save_name = f"{style_code}_{color_code}_IMAGEURL.jpg"
                        save_path = os.path.join(output_dir_color, save_name)
                        if os.path.exists(save_path):
                            print(f"款色图已存在，跳过: {save_name}")
                        else:
                            img_color = extract_first_image_from_folder(color_path)
                            if img_color:
                                img_color.save(save_path)
                                print(f"保存款色图: {save_name}")
                            else:
                                print(f"款色文件夹无图片: {color_path}")
                    else:
                        print(f"⚠️ 未匹配款色：{style_code} / {color_desc}")

    print(f"\n🎉 全部处理完成，款号图保存在：{output_dir_style}，款色图保存在：{output_dir_color}")

if __name__ == "__main__":
    main()
