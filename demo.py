"""
模块名称: main.py

功能：
- 实现一个随机选择人员的程序
- 提供导入名单、导出名单、重置名单等功能

作者: xhd0728
日期: 2023年6月22日
网站: https://github.com/xhd0728/simple-sample
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Scrollbar
import pandas as pd
import ctypes


class App:
    """
    类描述：随机选择人员的应用程序类

    属性：
    - root: Tkinter根窗口
    - screen_width: 屏幕宽度
    - screen_height: 屏幕高度
    - menu_bar: 应用程序的菜单栏
    - file_menu: 文件菜单
    - start_button: 开始按钮
    - save_button: 记录按钮
    - revoke_button: 撤销按钮
    - is_scrolling: 是否正在滚动选择人员
    - teacher_data: 导入的教师名单数据
    - teacher_num: 教师总数
    - export_data: 导出的名单数据
    - export_num: 已记录的人员数量
    - config_label: 用于显示人员总数、已记录数和剩余数的标签
    - label: 用于显示随机选择的人员信息的标签
    - result: 显示已选记录的文本框
    - yscrollbar: 已选记录文本框的纵向滚动条
    """

    def __init__(self, root):
        """
        初始化应用程序

        参数：
        - root: Tkinter根窗口
        """
        self.root = root
        user32 = ctypes.windll.user32
        self.screen_width = user32.GetSystemMetrics(0)
        self.screen_height = user32.GetSystemMetrics(1)
        pad_width = round((self.screen_width-800)/2)
        pad_height = round((self.screen_height-480)/2)
        self.root.geometry(f"800x500+{pad_width}+{pad_height}")
        self.root.title("Random")

        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_command(label="导入名单", command=self.import_list)
        self.menu_bar.add_command(label="导出名单", command=self.export_list)
        self.menu_bar.add_command(label="重置所有", command=self.reset_list)

        self.menu_bar.add_command(label="关于", command=self.show_about)

        self.start_button = ttk.Button(self.root,
                                       text="开始",
                                       command=self.start_scrolling)
        self.start_button.place(x=150, y=400)

        self.save_button = ttk.Button(self.root,
                                      text="记录",
                                      command=self.save_data)
        self.save_button.place(x=250, y=400)

        self.revoke_button = ttk.Button(self.root,
                                        text="撤销",
                                        command=self.revoke_save)
        self.revoke_button.place(x=350, y=400)

        self.is_scrolling = False
        self.teacher_data = pd.DataFrame(columns=["工号", "姓名"])
        self.teacher_num = 0
        self.export_data = pd.DataFrame(columns=["序号", "工号", "姓名"])
        self.export_num = 0

        self.config_label = tk.Label(self.root,
                                     text=f"人员总数: {self.teacher_num} 已记录: {self.export_num} 剩余: {self.teacher_num-self.export_num}",
                                     font=("Times New Roman", 16),
                                     justify="center")
        self.config_label.place(x=100, y=20)

        self.label = tk.Label(self.root,
                              text="",
                              font=("Times New Roman", 36),
                              justify="center")
        self.label.place(x=30, y=150)

        self.result = tk.Text(self.root,
                              height=22,
                              width=24,
                              font=("Times New Roman", 12))
        self.yscrollbar = Scrollbar(self.root)
        self.yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result.place(x=550, y=0)
        self.yscrollbar.config(command=self.result.yview)
        self.result.config(yscrollcommand=self.yscrollbar.set)
        self.result.insert(tk.END, "———已选记录 ———\n")

    def start_scrolling(self):
        """
        开始/停止滚动选择人员
        """
        if self.teacher_data.empty:
            messagebox.showwarning("错误", "请先导入人员名单")
            return
        if self.is_scrolling:
            self.is_scrolling = False
            self.start_button.config(text="开始")
        else:
            self.is_scrolling = True
            self.start_button.config(text="停止")
            self.scroll_names()

    def scroll_names(self):
        """
        滚动显示随机选择的人员信息
        """
        if not self.is_scrolling:
            return
        row = self.get_random_row()
        teacher_id = str(row["工号"].values[0])
        teacher_name = str(row["姓名"].values[0])
        if len(teacher_id) == 9:
            teacher_id = '0'+teacher_id
        self.label.config(text=f'{teacher_id} {teacher_name}')
        self.root.after(100, self.scroll_names)

    def save_data(self):
        """
        记录已选人员信息
        """
        if self.is_scrolling:
            self.is_scrolling = False
            self.start_button.config(text="开始")
        if not self.label["text"]:
            messagebox.showwarning("错误", "未选择人员")
            return
        row = str(self.label["text"]).split()
        if len(row) < 2:
            messagebox.showwarning("错误", "添加失败")
            return
        if row[0] in self.export_data.values:
            messagebox.showwarning("错误", "人员已记录")
            return
        self.export_data.loc[len(self.export_data)] = [self.export_num+1] + row
        self.export_num += 1
        self.result.insert(tk.END, f'{self.export_num} {row[0]} {row[1]}\n')
        self.result.focus_force()
        self.result.see(tk.END)
        self.update_config()

    def revoke_save(self):
        """
        撤销最近一次记录
        """
        if not self.export_num:
            messagebox.showwarning("错误", "无历史记录")
            return
        totalLen = len(self.result.get(1.0, tk.END).split("\n"))
        delstart = f"{totalLen-2}.0"
        delend = f"{totalLen}.0"
        self.result.delete(delstart, delend)
        self.export_data = self.export_data[:-1]
        self.export_num -= 1
        self.result.focus_force()
        self.result.insert(tk.END, '\n')
        self.update_config()

    def import_list(self):
        """
        导入人员名单
        """
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            try:
                data = pd.read_excel(file_path)
                if not self.valid_data(data):
                    messagebox.showwarning("错误", "无法解析Excel文件\n请选择正确的Excel文件")
                    return
                self.teacher_data = data
                self.teacher_num = len(data)
                self.update_config()
                messagebox.showinfo("上传成功",
                                    f"成功导入{self.teacher_num}条数据")
            except pd.errors.ParserError:
                messagebox.showwarning("错误", "无法解析文件\n请选择正确的Excel文件")

    def reset_list(self):
        """
        重置人员名单和记录信息
        """
        self.teacher_data = pd.DataFrame(columns=["工号", "姓名"])
        self.export_data = pd.DataFrame(columns=["序号", "工号", "姓名"])
        self.teacher_num = 0
        self.export_num = 0
        self.label.config(text="")
        self.result.delete('1.0', 'end')
        self.result.insert(tk.END, "———已选记录 ———\n")
        self.update_config()
        messagebox.showinfo("成功", "重置成功")

    def export_list(self):
        """
        导出已记录的人员信息
        """
        if self.export_data.empty:
            messagebox.showwarning("错误", "名单为空")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.export_data.to_excel(file_path, index=False)
            messagebox.showinfo("保存成功",
                                f"成功导出{self.export_num}条数据\n路径: {file_path}")

    def show_about(self):
        """
        显示关于窗口
        """
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        pad_width = round((self.screen_width-350)/2)
        pad_height = round((self.screen_height-150)/2)
        about_window.geometry(f"350x150+{pad_width}+{pad_height}")

        label = tk.Label(
            about_window,
            text="@Project: Random\n"
                 "@Author: xhd0728\n"
                 "@Email: xhd0728@hrbeu.edu.cn",
            font=("Times New Roman", 12),
            anchor="w",
            justify="left")
        label.pack(pady=20, padx=40)

    def get_random_row(self) -> pd.DataFrame:
        """
        从人员数据中随机选取一行
        """
        return self.teacher_data.sample(n=1)

    def update_config(self):
        """
        更新配置信息
        """
        self.config_label.config(
            text=f"人员总数: {self.teacher_num} 已记录: {self.export_num} 剩余: {self.teacher_num-self.export_num}")

    def valid_data(self, df) -> bool:
        """
        验证数据是否合法
        """
        if '工号' in df.columns and '姓名' in df.columns:
            return True
        return False


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.resizable(0, 0)
    root.mainloop()
