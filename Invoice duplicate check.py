import os
import fitz  # PyMuPDF
import cv2
import pytesseract
import numpy as np
import re
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askdirectory

# 设置Tesseract路径
pytesseract.pytesseract.tesseract_cmd = r'D:\software\Tesseract-OCR\tesseract.exe'


def select_roi(img):
    """手动选择识别区域"""
    cv2.namedWindow("选择区域", cv2.WINDOW_NORMAL)
    clone = img.copy()

    roi = []

    def click_and_crop(event, x, y, flags, param):
        nonlocal roi
        if event == cv2.EVENT_LBUTTONDOWN:
            roi = [(x, y)]
        elif event == cv2.EVENT_LBUTTONUP:
            roi.append((x, y))
            cv2.rectangle(clone, roi[0], roi[1], (0, 255, 0), 2)
            cv2.imshow("选择区域", clone)

    cv2.setMouseCallback("选择区域", click_and_crop)

    while True:
        cv2.imshow("选择区域", clone)
        key = cv2.waitKey(1) & 0xFF
        if key == ord("q"):  # 按 'q' 键退出
            break

    cv2.destroyAllWindows()

    if len(roi) == 2:
        return (roi[0][0], roi[0][1], roi[1][0], roi[1][1])
    return None


def extract_invoice_info(pdf_path):
    """从PDF中提取发票信息"""
    doc = fitz.open(pdf_path)
    invoice_data = []

    for page_number, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 提高分辨率
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, -1)

        # 手动选择区域
        roi = select_roi(img)
        if roi:
            x1, y1, x2, y2 = roi
            cropped_img = img[y1:y2, x1:x2]  # 截取选定区域

            # 保存选定区域图像
            output_image_path = f"selected_region_page_{page_number + 1}.png"
            cv2.imwrite(output_image_path, cropped_img)

            # 使用OCR提取文本，指定中文语言包
            text = pytesseract.image_to_string(cropped_img, lang='chi_sim')

            # 只提取20位数字
            twenty_digit_numbers = re.findall(r'\b\d{20}\b', text)
            for number in twenty_digit_numbers:
                invoice_data.append({'文件名': os.path.basename(pdf_path), '20位数字': number})  # 添加文件名和数字

    return invoice_data


def select_directory():
    """选择文件夹"""
    Tk().withdraw()  # 隐藏主窗口
    folder_path = askdirectory(title="选择PDF存储目录")
    return folder_path


# 选择PDF文件夹路径
pdf_folder = select_directory()
if pdf_folder:
    all_invoice_data = []

    for filename in os.listdir(pdf_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(pdf_folder, filename)
            invoice_info = extract_invoice_info(pdf_path)
            all_invoice_data.extend(invoice_info)

    # 将数据转换为DataFrame
    df = pd.DataFrame(all_invoice_data)

    # 保存到Excel文件
    output_file = os.path.join(pdf_folder, '发票信息.xlsx')
    df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"发票信息已保存到 {output_file}")
else:
    print("未选择任何目录。")
