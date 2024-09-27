import os
import pandas as pd
from datetime import datetime
from PIL import Image, ExifTags
import pillow_heif


def heic_to_jpg(input_file):
    heif_file = pillow_heif.read_heif(input_file)
    image = Image.frombytes(
        heif_file.mode,
        heif_file.size,
        heif_file.data,
        "raw",
        heif_file.mode,
        heif_file.stride,
    )
    return image


def get_image_orientation(image_path):
    if image_path.endswith(".HEIC"):
        return None
    img = Image.open(image_path)

    # 检查是否包含 EXIF 信息
    exif_data = img._getexif()
    if exif_data is None:
        return None

    # 查找 EXIF 的 Orientation 标签编号
    for tag, value in ExifTags.TAGS.items():
        if value == 'Orientation':
            orientation_tag = tag
            break
    else:
        return None

    # 获取 Orientation 信息
    orientation = exif_data.get(orientation_tag)

    # 根据 Orientation 值返回对应的旋转角度
    if orientation == 1:
        return 0      # 正常 (No rotation)
    elif orientation in [3, 2]:
        return 180    # 旋转180度
    elif orientation == 6:
        return 270    # 顺时针旋转90度
    elif orientation == 8:
        return 90     # 逆时针旋转90度
    else:
        return None


def get_images_sorted_by_creation_time(folder_path):
    all_files = os.listdir(folder_path)
    image_files = [f for f in all_files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', 'heic'))]
    image_files_with_time = [(f, os.path.getctime(os.path.join(folder_path, f))) for f in image_files]
    # sorted_files = sorted(image_files_with_time, key=lambda x: x[1])
    sorted_files_with_time = [(f, datetime.fromtimestamp(ctime).strftime('%Y-%m-%d %H:%M:%S')) for f, ctime in image_files_with_time]

    return sorted_files_with_time


def read_excel_file(file_path, sheet_name=None, header_row=0):
    if sheet_name:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        # 提取指定的列
        if id_column_name in df.columns and name_column_name in df.columns:
            id_name_list = list(zip(df[id_column_name], df[name_column_name]))
            return id_name_list
    else:
        df = pd.read_excel(file_path, sheet_name=None)
        print("读取所有工作表的数据:")

    return df


def rotate_image(image_path):
    if image_path.endswith(".HEIC"):
        img = heic_to_jpg(image_path)
    else:
        # 打开图像文件
        img = Image.open(image_path).convert('RGB')

    # 获取旋转角度
    rotation_angle = get_image_orientation(image_path)

    # 如果有旋转角度，执行旋转操作
    if rotation_angle and rotation_angle != 0:
        print(f"旋转图像 {rotation_angle} 度")
        # 使用 Pillow 的 rotate 方法进行旋转（expand=True 确保图像不会被裁剪）
        img = img.rotate(angle=rotation_angle, expand=True)

    return img


def images_to_pdf(image_files_list, output_pdf):
    images = []
    for img in image_files_list:
        angel = get_image_orientation(os.path.join(images_dir, img))
        images.append(rotate_image(os.path.join(images_dir, img)))

    if images:
        images[0].save(output_pdf, save_all=True, append_images=images[1:])
        print(f"PDF 文件已保存为: {output_pdf}")
    else:
        print("没有提供有效的图像文件。")


if __name__ == "__main__":
    global images_dir, excel_path, id_column_name, name_column_name, mode
    mode = "点名册"     # "点名册"
    id_column_name = "学号"
    name_column_name = "姓名"
    images_dir = "Image"
    excel_path = "Excel/名单.xlsx"

    image_files = get_images_sorted_by_creation_time(images_dir)
    print("图片数量：", len(image_files))

    if mode == "点名册":
        output_pdf = os.path.join(images_dir, mode + ".pdf")
        item_image_files = [x[0] for x in image_files]
        images_to_pdf(item_image_files, output_pdf)
        exit()

    elif mode == '个人成绩':
        data = read_excel_file(excel_path, sheet_name="Sheet1")
        print("人头数：", len(data))

        i = 0
        for item in data:
            item_image_files = [x[0] for x in image_files[i:i+8]]
            output_pdf = os.path.join(images_dir, str(item[0]) + '_' + item[1] + ".pdf")
            images_to_pdf(item_image_files, output_pdf)
            i += 8
    else:
        print("Error! mode must be in ['点名册', '个人成绩']")

