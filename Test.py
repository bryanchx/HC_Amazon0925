import itchat
import tkinter as tk
from tkinter import scrolledtext

class WeChatApp:
    def __init__(self, master):
        self.master = master
        master.title("微信聊天")

        self.text_area = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=50, height=20)
        self.text_area.pack(padx=10, pady=10)

        self.entry = tk.Entry(master, width=50)
        self.entry.pack(padx=10, pady=10)
        self.entry.bind("<Return>", self.send_message)

        self.send_button = tk.Button(master, text="发送", command=self.send_message)
        self.send_button.pack(padx=10, pady=10)

        self.text_area.insert(tk.END, "登录微信...\n")
        itchat.auto_login(hotReload=False)  # 设置为 False
        itchat.run(blockThread=False)
        itchat.msg_register(itchat.content.TEXT)(self.receive_message)

    def send_message(self, event=None):
        message = self.entry.get()
        if message:
            self.text_area.insert(tk.END, f"发送: {message}\n")
            itchat.send(message, toUserName='filehelper')
            self.entry.delete(0, tk.END)

    def receive_message(self, msg):
        self.text_area.insert(tk.END, f"收到: {msg.text}\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = WeChatApp(root)
    root.mainloop()
