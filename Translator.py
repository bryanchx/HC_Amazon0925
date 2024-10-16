import pandas as pd
from googletrans import Translator
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tkinter import ttk

# 创建 Tkinter 应用程序
root = tk.Tk()
root.withdraw()  # 隐藏主窗口

# 文件选择对话框
file_path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel files", "*.xlsx;*.xls")])
if not file_path:
    messagebox.showerror("错误", "未选择文件！")
    exit()

# 读取工作表名称
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

# 选择工作表
sheet_name = simpledialog.askstring("输入", "请选择工作表：\n" + "\n".join(sheet_names))
if sheet_name not in sheet_names:
    messagebox.showerror("错误", "工作表不存在！")
    exit()

# 读取指定的工作表
df = pd.read_excel(file_path, sheet_name=sheet_name)

# 创建翻译器对象
translator = Translator()

# 手动选择翻译列
column_names = df.columns.tolist()
column_selection = simpledialog.askstring("输入", "请选择需要翻译的列（默认是 A 列，输入列名）：", initialvalue="A")

# 默认情况下如果输入为空，则使用 A 列
source_column = column_selection if column_selection in column_names else column_names[0]

# 指定中文翻译列名
target_column = '中文翻译'

# 定义翻译函数
def translate_text(text):
    try:
        translated = translator.translate(text, src='en', dest='zh-cn')
        return translated.text
    except Exception as e:
        print(f"翻译失败: {e}")
        return text  # 返回原文本

# 在新列中存放中文翻译
df[target_column] = df[source_column].apply(translate_text)

# 保存到新的 Excel 文件
output_file_path = filedialog.asksaveasfilename(title="保存翻译文件", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls")])
if output_file_path:
    df.to_excel(output_file_path, index=False)
    messagebox.showinfo("成功", "翻译完成，文件已保存。")
else:
    messagebox.showwarning("警告", "未保存文件！")
