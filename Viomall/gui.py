import tkinter as tk
from tkinter import ttk, messagebox
from config import USERNAME, PASSWORD
from api_client import APIClient


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("API Client")
        self.root.geometry("800x600")
        self.api_client = APIClient()

        self.menubar = tk.Menu(self.root)
        self.create_menu()

        self.login_frame = LoginFrame(root, self)
        self.login_frame.pack(padx=10, pady=10)

        self.main_frame = MainFrame(root, self)

    def create_menu(self):
        product_menu = tk.Menu(self.menubar, tearoff=0)
        sub_menu = tk.Menu(product_menu, tearoff=0)
        sub_menu.add_command(label="待处理", command=self.show_pending_frame)
        sub_menu.add_command(label="已完成", command=self.show_completed_frame)

        product_menu.add_cascade(label="Amazon刊登中转站", menu=sub_menu)
        product_menu.add_command(label="赏金猎人", command=self.show_hunter_frame)
        product_menu.add_command(label="挖宝促销推荐", command=self.show_promotion_frame)

        self.menubar.add_cascade(label="选品", menu=product_menu)
        self.menubar.add_cascade(label="销售", menu=tk.Menu(self.menubar, tearoff=0))
        self.menubar.add_cascade(label="订单", menu=tk.Menu(self.menubar, tearoff=0))
        self.menubar.add_cascade(label="营销", menu=tk.Menu(self.menubar, tearoff=0))
        self.menubar.add_cascade(label="社区", menu=tk.Menu(self.menubar, tearoff=0))

    def show_main_frame(self):
        self.login_frame.pack_forget()
        self.main_frame.pack(padx=10, pady=10)
        self.show_menu()

    def show_pending_frame(self):
        self.main_frame.display_data([])
        self.main_frame.call_api("获取待处理列表")

    def show_completed_frame(self):
        self.main_frame.display_data([])
        self.main_frame.call_api("获取已完成列表")

    def show_hunter_frame(self):
        messagebox.showinfo("功能未实现", "赏金猎人功能未实现！")

    def show_promotion_frame(self):
        messagebox.showinfo("功能未实现", "挖宝促销推荐功能未实现！")

    def show_menu(self):
        self.root.config(menu=self.menubar)


class LoginFrame(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="用户名", font=("Arial", 12)).grid(row=0, column=0, pady=(10, 5))
        self.username_entry = tk.Entry(self, font=("Arial", 12))
        self.username_entry.insert(0, USERNAME)
        self.username_entry.grid(row=0, column=1, pady=(10, 5))

        tk.Label(self, text="密码", font=("Arial", 12)).grid(row=1, column=0)
        self.password_entry = tk.Entry(self, show='*', font=("Arial", 12))
        self.password_entry.insert(0, PASSWORD)
        self.password_entry.grid(row=1, column=1)

        self.remember_password_var = tk.BooleanVar()
        tk.Checkbutton(self, text="记住密码", variable=self.remember_password_var, font=("Arial", 12)).grid(row=2,
                                                                                                            columnspan=2,
                                                                                                            pady=(
                                                                                                            5, 10))

        tk.Button(self, text="登录", command=self.login, font=("Arial", 12)).grid(row=3, columnspan=2, pady=(5, 10))

        self.toast_label = tk.Label(self, text="", font=("Arial", 12), fg="green")
        self.toast_label.grid(row=4, columnspan=2)

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if self.app.api_client.login(username, password):
            if self.remember_password_var.get():
                print("已选择记住密码")
                # TODO: 保存用户名和密码的逻辑
            self.toast_label.config(text="登录成功！")
            self.after(2000, self.clear_toast)
            self.app.show_main_frame()
        else:
            self.toast_label.config(text="登录失败！", fg="red")
            self.after(2000, self.clear_toast)

    def clear_toast(self):
        self.toast_label.config(text="")


class MainFrame(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.current_page = 1
        self.total_pages = 1
        self.create_widgets()

    def create_widgets(self):
        self.data_listbox = tk.Listbox(self, selectmode=tk.MULTIPLE, width=70, height=15, font=("Arial", 12))
        self.data_listbox.pack(padx=10, pady=10)

        self.scrollbar = ttk.Scrollbar(self)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.data_listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.data_listbox.yview)

        self.loading_label = tk.Label(self, text="", font=("Arial", 12), fg="blue")
        self.loading_label.pack(pady=(0, 10))

        self.view_button = tk.Button(self, text="查看", command=self.view_selected_items, font=("Arial", 12))  # 查看按钮
        self.view_button.pack(pady=(0, 10))

        # 翻页功能
        self.page_frame = tk.Frame(self)
        self.page_frame.pack(pady=(0, 10))

        self.prev_button = tk.Button(self.page_frame, text="<< 上一页", command=self.previous_page, font=("Arial", 12))
        self.prev_button.pack(side=tk.LEFT)

        self.page_label = tk.Label(self.page_frame, text=f"第 {self.current_page} 页 / {self.total_pages} 页",
                                   font=("Arial", 12))
        self.page_label.pack(side=tk.LEFT, padx=(10, 10))

        self.next_button = tk.Button(self.page_frame, text="下一页 >>", command=self.next_page, font=("Arial", 12))
        self.next_button.pack(side=tk.LEFT)

    def call_api(self, selected_api):
        self.loading_label.config(text="加载中...")
        self.data_listbox.delete(0, tk.END)

        if selected_api == "获取已完成列表":
            complete_api_list = self.app.api_client.get_amazon_listing(self.current_page, 10)
            if complete_api_list and complete_api_list.get('success'):
                self.display_data(complete_api_list['value']['data'])
                self.total_pages = (complete_api_list['value']['pagination']['total'] + 9) // 10
                self.page_label.config(text=f"第 {self.current_page} 页 / {self.total_pages} 页")
            else:
                messagebox.showerror("接口调用", "获取已完成列表失败！")

        elif selected_api == "获取待处理列表":
            # 获取待处理列表的逻辑（待实现）
            messagebox.showinfo("接口调用", "获取待处理列表成功！")

        self.loading_label.config(text="")

    def display_data(self, data):
        self.data_listbox.delete(0, tk.END)
        for item in data:
            self.data_listbox.insert(tk.END, f"ID: {item['id']}, 标题: {item['title']}")

    def view_selected_items(self):
        selected_indices = self.data_listbox.curselection()  # 获取选中的索引
        selected_items = [self.data_listbox.get(i) for i in selected_indices]  # 获取选中的项

        if not selected_items:
            messagebox.showinfo("信息", "未选择任何项！")
            return

        for item in selected_items:
            item_id = item.split(",")[0].split(":")[1].strip()  # 从列表项中提取ID
            self.loading_label.config(text=f"加载中... 正在获取 ID: {item_id}")  # 更新加载提示
            product_info = self.app.api_client.get_amazon_listing_product_list(item_id)  # 调用获取详细信息的API

            if product_info and product_info.get('success'):
                messagebox.showinfo("产品详情", f"ID: {item_id}, 详情: {product_info['value']}")
            else:
                messagebox.showerror("接口调用", f"获取 ID: {item_id} 的详情失败！")

        self.loading_label.config(text="")  # 隐藏加载提示

    def previous_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.call_api("获取已完成列表" if self.current_page == 1 else "获取待处理列表")

    def next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.call_api("获取已完成列表" if self.current_page == 1 else "获取待处理列表")

if __name__ == "__main__":
    root = tk.Tk()  # 创建主窗口
    app = App(root)  # 创建应用实例
    root.mainloop()  # 进入主循环
