import pandas as pd
from tkinter import Tk, filedialog
import re

# 读取TXT文件
def process_txt(output_file):
    # 打开文件选择对话框
    Tk().withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(title="Select TXT File", filetypes=[("Text Files", "*.txt")])

    if not file_path:
        print("No file selected. Exiting.")
        return

    # 读取TXT文件，假设以制表符（\t）分隔
    df = pd.read_csv(file_path, sep='\t')

    # 确保A列是SKU，B列是ASIN
    if 'sku' not in df.columns or 'asin' not in df.columns:
        raise ValueError("The TXT file must contain 'SKU' and 'ASIN' columns.")

    # 提取日期范围
    def extract_date(sku):
        match = re.search(r'CHX(\d{8})-', sku)
        return match.group(1) if match else None

    df['Date'] = df['sku'].apply(extract_date)

    # 获取B列（ASIN）和日期
    asins = df[['asin', 'Date']].dropna()

    # 按日期分组并每2900个拼接一次
    result = []
    for date, group in asins.groupby('Date'):
        asin_list = group['asin'].tolist()
        chunks = [';'.join(asin_list[i:i+2900]) for i in range(0, len(asin_list), 2900)]
        for chunk in chunks:
            result.append({'Date': date, 'Chunk': chunk})

    # 创建一个新DataFrame保存结果
    result_df = pd.DataFrame(result)

    # 将结果写入新的Excel文件
    result_df.to_excel(output_file, index=False)
    print(f"Processed ASIN chunks with dates have been saved to {output_file}")

# 调用函数
# 输出文件路径
output_file = 'output_chunks_with_dates.xlsx'  # 替换为你的输出文件路径

process_txt(output_file)
