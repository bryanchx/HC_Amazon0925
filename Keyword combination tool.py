import pandas as pd
from itertools import product
import os
from tkinter import Tk, Label, Button, StringVar, Listbox, Scrollbar, Checkbutton, IntVar

from tkinter.filedialog import askopenfilename  # 确保导入这一行


def keyword_combination(keywords):
    combinations = []
    for combo in product(*keywords):
        combinations.append(' '.join(combo))  # 组合成标题
    return combinations


def capitalize_keywords(text, capitalize_prepositions):
    small_words = {'and', 'or', 'but', 'in', 'of', 'the', 'a', 'an', 'for', 'to', 'with', 'on', 'at', 'as', 'by',
                   'from'}

    words = text.split()
    if not capitalize_prepositions:
        return ' '.join(word.lower() if word.lower() in small_words else word.capitalize() for word in words)

    return ' '.join(word.capitalize() for word in words)


def select_file():
    file_path = askopenfilename(title="选择Excel文件", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        file_var.set(file_path)
        load_sheets(file_path)


def load_sheets(file_path):
    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names
    sheet_list.delete(0, 'end')
    for sheet in sheets:
        sheet_list.insert('end', sheet)


def generate_combinations():
    selected_sheet = sheet_list.get(sheet_list.curselection())
    file_path = file_var.get()
    df = pd.read_excel(file_path, sheet_name=selected_sheet)

    # 将每列转换为列表
    keyword_groups = [df[col].dropna().tolist() for col in df.columns[:7]]  # 只取前7列

    # 生成组合
    combinations = keyword_combination(keyword_groups)

    # 处理标题，除介词外首字母大写
    processed_combinations = [capitalize_keywords(combo, not capitalize_var.get()) for combo in combinations]

    # 将组合写入到同一个Excel的新Sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        output_df = pd.DataFrame(processed_combinations, columns=['Keyword Combination'])
        output_df.to_excel(writer, sheet_name='Keyword Combinations', index=False)

    Label(root, text=f'组合已成功写入到 "{os.path.basename(file_path)}" 的新Sheet中.').pack()


# 创建Tkinter窗口
root = Tk()
root.title("关键词组合生成器")

# 文件选择部分
file_var = StringVar()
Label(root, text="选择Excel文件:").pack()
Button(root, text="浏览", command=select_file).pack()
Label(root, textvariable=file_var).pack()

# 工作表选择部分
Label(root, text="选择工作表:").pack()
sheet_list = Listbox(root, width=50, height=10)
sheet_list.pack()

# 勾选框部分
capitalize_var = IntVar(value=1)  # 默认勾选
Checkbutton(root, text="首字母大写，除介词外", variable=capitalize_var).pack()

# 生成组合按钮
Button(root, text="生成组合", command=generate_combinations).pack()

root.mainloop()
