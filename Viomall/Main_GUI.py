import os
import time
import random
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import requests
import json
import sqlite3


# 登录
def login(username, password):
    login_url = "https://dms.viomall.com/v3/user/userLogin.do"
    session = requests.Session()

    payload = f"userName={username}&passWord={password}"
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Origin': 'https://dms.viomall.com',
        'Referer': 'https://dms.viomall.com/v3/users/login.do',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
    }

    # 发送登录请求
    login_response = session.post(login_url, headers=headers, data=payload)

    # 检查登录是否成功
    if login_response.ok:
        return session  # 返回会话对象以供后续请求使用
    else:
        return None

# 获取Api刊登工具已完成列表数据
def get_amazon_listing(session, page, pageSize):
    url = "https://dms.viomall.com/v3/amazonListing/getAmazonListingBatchList.do"
    payload = f'display=completed&pageSize={pageSize}&page={page}&sortField=&sortOrder='
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://dms.viomall.com',
        'Referer': 'https://dms.viomall.com/v3/amazonListing/amazonListingBatchList.do',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
    }

    # 发送请求
    response = session.post(url, headers=headers, data=payload)
    if response.ok:
        return response.json()
    else:
        return None

# 获取 Api刊登工具里面的SKU列表
def get_amazon_listing_product_list(session, product_id, page=1, page_size=10):
    url = f"https://dms.viomall.com/v3/amazonListing/getAmazonListingBatchProductList.do?id={product_id}"
    payload = f'display=all&page={page}&pageSize={page_size}&sortField=&sortOrder='
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://dms.viomall.com',
        'Referer': f'https://dms.viomall.com/v3/amazonListing/editAmazonListingBatch.do?id={product_id}',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
    }

    # 发送请求
    response = session.post(url, headers=headers, data=payload)

    if response.ok:
        return response.json()
    else:
        return None

# 保存 SKU 到 Excel
def save_skus_to_excel(skus):
    df = pd.DataFrame(skus, columns=['productSku'])
    excel_file = 'product_skus.xlsx'

    # 检查文件是否存在
    if os.path.exists(excel_file):
        df_existing = pd.read_excel(excel_file, sheet_name='ProductSKUs')
        df_combined = pd.concat([df_existing, df], ignore_index=True)
    else:
        df_combined = df

    # 保存到 Excel
    with pd.ExcelWriter(excel_file, mode='w', engine='openpyxl') as writer:
        df_combined.to_excel(writer, index=False, header=True, sheet_name='ProductSKUs')

    print("SKU已成功保存到Excel文件！")

# 初始化数据库，保存登录信息
def init_db():
    conn = sqlite3.connect("user_settings.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            password TEXT,
            remember INTEGER
        )
    ''')
    conn.commit()
    conn.close()

# 保存用户设置到数据库
def save_user_settings(username, password, remember):
    conn = sqlite3.connect("user_settings.db")
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO settings (username, password, remember) 
        VALUES (?, ?, ?)
    ''', (username, password, remember))
    conn.commit()
    conn.close()

# 获取已保存的用户设置
def get_user_settings():
    conn = sqlite3.connect("user_settings.db")
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM settings ORDER BY id DESC LIMIT 1')
    result = cursor.fetchone()
    conn.close()
    return result

# 登录窗口界面
def login_window():
    def on_login():
        username = username_entry.get()
        password = password_entry.get()
        remember = remember_var.get()

        if not username or not password:
            messagebox.showerror("错误", "用户名和密码不能为空！")
            return

        session = login(username, password)

        if session:
            if remember:
                save_user_settings(username, password, remember)
            open_main_window(session)
        else:
            messagebox.showerror("登录失败", "用户名或密码错误！")

    # 创建登录窗口
    login_win = tk.Tk()
    login_win.title("登录")
    login_win.geometry("400x300")

    # 用户名输入框
    tk.Label(login_win, text="用户名:").pack(pady=10)
    username_entry = tk.Entry(login_win, width=30)
    username_entry.pack()

    # 密码输入框
    tk.Label(login_win, text="密码:").pack(pady=10)
    password_entry = tk.Entry(login_win, width=30, show="*")
    password_entry.pack()

    # 记住密码勾选框
    remember_var = tk.IntVar()
    remember_check = tk.Checkbutton(login_win, text="记住密码", variable=remember_var)
    remember_check.pack()

    # 登录按钮
    login_button = tk.Button(login_win, text="登录", command=on_login)
    login_button.pack(pady=20)

    # 获取保存的设置（如果有）
    settings = get_user_settings()
    if settings:
        username_entry.insert(0, settings[1])
        password_entry.insert(0, settings[2])
        remember_var.set(settings[3])

    login_win.mainloop()

# 主界面
def open_main_window(session):
    def fetch_skus():
        complete_api_list = get_amazon_listing(session, 1, 150)
        if complete_api_list and complete_api_list['success']:
            listings = complete_api_list['value']['data']
            all_skus = []
            for listing in listings:
                item_id = listing['id']
                listing_product_list = get_amazon_listing_product_list(session, item_id, 1, 1000)
                if listing_product_list and listing_product_list['success']:
                    product_skus = [item['productSku'] for item in listing_product_list['value']['data']]
                    all_skus.extend(product_skus)

            if all_skus:
                save_skus_to_excel(all_skus)
                messagebox.showinfo("成功", "SKU已成功保存到Excel文件！")
            else:
                messagebox.showwarning("警告", "没有获取到任何SKU！")
        else:
            messagebox.showerror("错误", "无法获取产品列表！")

        # 创建主界面窗口
        main_win = tk.Tk()
        main_win.title("主界面")
        main_win.geometry("500x400")

        # 菜单栏
        menu_bar = tk.Menu(main_win)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="退出", command=main_win.quit)
        menu_bar.add_cascade(label="文件", menu=file_menu)

        main_win.config(menu=menu_bar)

        # 显示欢迎信息
        tk.Label(main_win, text="欢迎使用系统！", font=("Arial", 20)).pack(pady=20)

        # 获取产品列表并保存到 Excel 的按钮
        fetch_button = tk.Button(main_win, text="获取SKU并保存", width=30, height=2, command=fetch_skus)
        fetch_button.pack(pady=20)

        # 退出按钮
        exit_button = tk.Button(main_win, text="退出", width=30, height=2, command=main_win.quit)
        exit_button.pack(pady=10)

        # 运行主界面
        main_win.mainloop()

# 初始化数据库
init_db()

# 启动登录窗口
login_window()

