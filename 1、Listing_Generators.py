import pandas as pd
import random
import os
from itertools import product
from tkinter import Tk, Label, Button, StringVar, Listbox, Checkbutton, IntVar, messagebox, Entry, Frame
from tkinter.filedialog import askopenfilename


def keyword_combination(keywords):
    combinations = []
    for combo in product(*keywords):
        combinations.append(' '.join(combo))
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
    if not file_path:
        messagebox.showwarning("警告", "文件路径为空，无法加载工作表！")
        return

    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names
    sheet_list.delete(0, 'end')
    for sheet in sheets:
        sheet_list.insert('end', sheet)


def generate_combinations():
    if not file_var.get():
        messagebox.showwarning("警告", "请先选择一个Excel文件！")
        return

    if sheet_list.curselection() == ():
        messagebox.showwarning("警告", "请先选择一个工作表！")
        return

    selected_sheet = sheet_list.get(sheet_list.curselection())
    file_path = file_var.get()
    df = pd.read_excel(file_path, sheet_name=selected_sheet)

    keyword_groups = [df[col].dropna().tolist() for col in df.columns[:7]]
    if not keyword_groups or all(not group for group in keyword_groups):
        messagebox.showwarning("警告", "关键词组为空，请检查Excel文件！")
        return

    combinations = keyword_combination(keyword_groups)
    processed_combinations = [capitalize_keywords(combo, not capitalize_var.get()) for combo in combinations]

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        output_df = pd.DataFrame(processed_combinations, columns=['Keyword Combination'])
        output_df.to_excel(writer, sheet_name='Keyword Combinations', index=False)

    messagebox.showinfo("成功", f'组合已成功写入到 "{os.path.basename(file_path)}" 的新Sheet中.')


def split_into_groups(data, target_length=465):
    if not data:
        return []

    groups = []
    current_group = []
    current_length = 0

    random.shuffle(data)

    for item in data:
        item_length = len(item)
        if current_length + item_length > target_length:
            if current_group:
                groups.append(current_group)
            current_group = [item]
            current_length = item_length
        else:
            current_group.append(item)
            current_length += item_length

    if current_group:
        groups.append(current_group)

    final_groups = []
    for group in groups:
        group_string = ' '.join(group)
        if len(group_string) > target_length:
            split_groups = []
            temp_group = []
            temp_length = 0
            for item in group:
                item_length = len(item)
                if temp_length + item_length > target_length:
                    split_groups.append(temp_group)
                    temp_group = [item]
                    temp_length = item_length
                else:
                    temp_group.append(item)
                    temp_length += item_length
            if temp_group:
                split_groups.append(temp_group)
            final_groups.extend(split_groups)
        else:
            final_groups.append(group)

    return final_groups


def generate_groups():
    if not file_var.get():
        messagebox.showwarning("警告", "请先选择一个Excel文件！")
        return

    file_path = file_var.get()
    df = pd.read_excel(file_path)

    if df.empty:
        messagebox.showwarning("警告", "Excel文件为空，请检查文件内容！")
        return

    data_column = df.iloc[:, 0].dropna().tolist()
    if not data_column:
        messagebox.showwarning("警告", "数据列为空，请检查Excel文件！")
        return

    try:
        num_runs = int(num_runs_entry.get())
        if num_runs <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("错误", "输入无效，请输入一个正整数。")
        return

    # 使用 with 语句来确保正确保存和关闭
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # 检查并创建 'Grouped Results' 工作表
        try:
            writer.book['Grouped Results']
        except KeyError:
            # 创建一个空的 DataFrame，并写入
            pd.DataFrame(columns=['Grouped Results']).to_excel(writer, sheet_name='Grouped Results', index=False)

        next_row = writer.sheets['Grouped Results'].max_row + 1  # 从最后一行开始写入

        for run in range(num_runs):
            random_data = data_column[:]
            random.shuffle(random_data)

            grouped_data = split_into_groups(random_data)
            for i, group in enumerate(grouped_data):
                group_string = ' '.join(group)
                writer.sheets['Grouped Results'].cell(row=next_row, column=i + 1, value=group_string)

            next_row += 1

    messagebox.showinfo("成功", f'分组结果已成功写入到 "{file_path}" 的 "Grouped Results" 中.')

# 创建Tkinter窗口
root = Tk()
root.title("关键词生成器")
root.geometry("800x600")
root.config(bg="#f7f7f7")

# 侧边导航栏
nav_frame = Frame(root, bg="#4CAF50", width=200)
nav_frame.pack(side="left", fill="y")

Label(nav_frame, text="功能导航", bg="#4CAF50", fg="white", font=("Arial", 16)).pack(pady=20)

Button(nav_frame, text="关键词组合", command=lambda: show_frame(combination_frame), bg="#4CAF50", fg="white").pack(
    pady=5)
Button(nav_frame, text="随机分组", command=lambda: show_frame(group_frame), bg="#4CAF50", fg="white").pack(pady=5)

# 主内容区
main_frame = Frame(root, bg="#ffffff")
main_frame.pack(side="right", fill="both", expand=True)

# 文件选择部分
file_var = StringVar()
file_frame = Frame(main_frame, bg="#ffffff", padx=20, pady=20)
file_frame.pack(pady=10)

Label(file_frame, text="选择Excel文件:", bg="#ffffff", font=("Arial", 12)).pack(pady=10)
Button(file_frame, text="浏览", command=select_file).pack(pady=5)

# 工作表选择部分
sheet_frame = Frame(main_frame, bg="#ffffff", padx=20, pady=20)
sheet_frame.pack(pady=10)

Label(sheet_frame, text="选择工作表:", bg="#ffffff", font=("Arial", 12)).pack(pady=10)
sheet_list = Listbox(sheet_frame, width=50, height=5)
sheet_list.pack(pady=5)

# 关键词组合部分
combination_frame = Frame(main_frame, bg="#ffffff", padx=20, pady=20)

capitalize_var = IntVar(value=1)
Checkbutton(combination_frame, text="首字母大写，除介词外", variable=capitalize_var, bg="#ffffff").pack(pady=5)
Button(combination_frame, text="生成关键词组合", command=generate_combinations).pack(pady=15)

# 随机分组部分
group_frame = Frame(main_frame, bg="#ffffff", padx=20, pady=20)

Label(group_frame, text="请输入要执行的次数:", bg="#ffffff", font=("Arial", 12)).pack(pady=10)
num_runs_entry = Entry(group_frame)
num_runs_entry.pack(pady=5)
Button(group_frame, text="生成随机分组", command=generate_groups).pack(pady=15)


# 显示特定框架的函数
def show_frame(frame):
    combination_frame.pack_forget()
    group_frame.pack_forget()
    frame.pack(pady=10)


# 初始化界面显示
show_frame(combination_frame)

root.mainloop()

