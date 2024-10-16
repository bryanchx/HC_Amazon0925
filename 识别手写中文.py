from paddleocr import PaddleOCR
import os
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

# 初始化 OCR 模型
ocr = PaddleOCR(use_angle_cls=True, lang='ch')  # lang='ch' 用于中文

# 读取图片
image_path = 'E:/workspace_py/HC_Amazon0925/jizhang.jpg'

# 执行 OCR 识别
result = ocr.ocr(image_path, cls=True)

# 只输出识别的中文
for line in result:
    for word_info in line:
        # word_info[1][0] 是识别的文本
        print(word_info[1][0])  # 只打印识别到的中文文本
