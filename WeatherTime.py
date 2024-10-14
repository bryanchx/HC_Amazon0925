import tkinter as tk
from tkinter import ttk
from datetime import datetime
import pytz
import requests
from bs4 import BeautifulSoup

# 根据城市获取时区
def get_timezone(city):
    timezones = {
        'Beijing': 'Asia/Shanghai',
        'Los Angeles': 'America/Los_Angeles',
        'London': 'Europe/London',
        'Tokyo': 'Asia/Tokyo',
        'Paris': 'Europe/Paris',
        'Sydney': 'Australia/Sydney'
    }
    return timezones.get(city, 'UTC')  # 默认时区为UTC

# 爬取天气信息
def get_weather(city):
    try:
        url = f"https://www.weather.com/zh-CN/weather/today/l/{city}"  # 示例URL，可能需要具体城市的地址
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # 解析天气信息（需要根据具体页面结构调整选择器）
        temperature = soup.find('span', class_='CurrentConditions--tempValue--3a50n').text
        weather_desc = soup.find('div', class_='CurrentConditions--phraseValue--2xXSr').text
        return f"{temperature}, {weather_desc}"
    except Exception as e:
        return "天气信息不可用"

# 更新所有城市的时间和天气
def update_time():
    for widget in result_frame.winfo_children():
        widget.destroy()  # 清空现有信息
    weather_info.delete(1.0, tk.END)  # 清空天气信息

    cities = city_entry.get().split(",")  # 用逗号分隔多个城市

    for city in cities:
        city = city.strip()  # 去除多余空格

        # 根据城市名设置对应的时区
        timezone = get_timezone(city)
        now = datetime.now(pytz.timezone(timezone))  # 获取当前时间

        # 创建新的卡片框架以放置城市信息
        city_frame = tk.Frame(result_frame, bg="#ffffff", bd=2, relief=tk.RAISED, padx=10, pady=10)
        city_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 城市名称标签
        city_label = tk.Label(city_frame, text=city, font=("Arial", 16, "bold"), bg="#ffffff")
        city_label.pack(side=tk.TOP, pady=(0, 5))

        # 时间标签
        time_label = tk.Label(city_frame, text=now.strftime("%Y-%m-%d %H:%M:%S"), font=("Arial", 14), bg="#ffffff")
        time_label.pack(side=tk.TOP, pady=(0, 5))

        # 更新标签的文本以显示最新时间
        time_label.after(1000, lambda t_label=time_label, tz=timezone: update_time_label(t_label, tz))

        # 获取天气信息并显示
        weather = get_weather(city)
        weather_info.insert(tk.END, f"{city}: {weather}\n")

# 更新时间标签的函数
def update_time_label(label, timezone):
    now = datetime.now(pytz.timezone(timezone))  # 获取当前时间
    label.config(text=now.strftime("%Y-%m-%d %H:%M:%S"))  # 更新标签文本
    label.after(1000, lambda: update_time_label(label, timezone))  # 每秒更新

# 创建主窗口
root = tk.Tk()
root.title("世界时间与天气应用")
root.geometry("800x500")
root.config(bg="#e0f7fa")

# 城市输入框
ttk.Label(root, text="输入城市名（用逗号分隔）:", background="#e0f7fa", font=("Arial", 14)).grid(column=0, row=0, padx=10, pady=10)
city_entry = ttk.Entry(root, width=50)
city_entry.grid(column=1, row=0, padx=10, pady=10)

# 设置默认城市
default_cities = "Beijing, Los Angeles, London"
city_entry.insert(0, default_cities)  # 插入默认城市

# 结果框架
result_frame = tk.Frame(root, bg="#e0f7fa")
result_frame.grid(column=0, row=1, columnspan=2, padx=10, pady=10)

# 天气信息框
weather_info = tk.Text(root, height=5, width=70, bg="#ffffff", font=("Arial", 12))
weather_info.grid(column=0, row=2, columnspan=2, padx=10, pady=10)
weather_info.insert(tk.END, "城市天气信息:\n")

# 更新按钮
update_button = ttk.Button(root, text="更新", command=update_time)
update_button.grid(column=0, row=3, columnspan=2, padx=10, pady=10)

# 启动应用
update_time()  # 启动时更新默认城市时间
root.mainloop()
