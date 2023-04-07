import fitz
import os


# 获取pdf文件的名称
def get_filename(file_path):
    path_list = os.path.split(file_path)
    file_list = os.path.splitext(path_list[1])
    return file_list[0]


# pdf 转 png 图片
def pdf_image(pdf_file, img_path, zoom_x=4, zoom_y=4, rotation_angle=0):
    pdf = fitz.open(pdf_file)
    for page_num in range(0, pdf.page_count):
        page_obj = pdf[page_num]
        # 创建用于图像变换的矩阵
        trans = fitz.Matrix(zoom_x, zoom_y)
        # 将PDF页面处理成图像
        pm = page_obj.get_pixmap(matrix=trans, alpha=False)
        temp = get_filename(pdf_file)
        # print("图片路径：" + f'{img_path}{temp}_{page_num + 1}.png')
        pm.save(f'{img_path}\\{temp}_{page_num + 1}.png')
    pdf.close()


# pdf_image("resources/GYJBoost.pdf", 'resources', 2, 2)
# pdf_image(pdf_path, filePath, 2, 2)
