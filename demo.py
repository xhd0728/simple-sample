"""
模块名称: demo.py

功能：
- 实现一个随机选择人员的程序
- 提供导入名单、导出名单、重置名单等功能

作者: xhd0728
日期: 2023年6月23日
网站: https://github.com/xhd0728/simple-sample
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Scrollbar
import pandas as pd
import ctypes

# 滚动时延/ms
SCROLL_INTERVAL_TIME = 50
# 教师工号长度
TEACHER_ID_LENGTH = 10
# 主窗口宽高/px
MAIN_WINDOW_WIDTH = 800
MAIN_WINDOW_HEIGHT = 500


class App:
    """
    类描述：随机选择人员的应用程序类

    属性：
    - root: Tkinter根窗口
    - screen_width: 屏幕宽度
    - screen_height: 屏幕高度
    - menu_bar: 应用程序的菜单栏
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
        self.root.protocol('WM_DELETE_WINDOW', self.window_close_handle)
        user32 = ctypes.windll.user32
        self.screen_width = user32.GetSystemMetrics(0)
        self.screen_height = user32.GetSystemMetrics(1)
        pad_width = round((self.screen_width-MAIN_WINDOW_WIDTH)/2)
        pad_height = round((self.screen_height-MAIN_WINDOW_HEIGHT)/2)
        self.root.geometry(
            f"{MAIN_WINDOW_WIDTH}x{MAIN_WINDOW_HEIGHT}+{pad_width}+{pad_height}")
        self.root.title("Random")

        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        self.menu_bar.add_command(label="导入名单", command=self.import_list)
        self.menu_bar.add_command(label="导出名单", command=self.export_list)
        self.menu_bar.add_command(label="重置所有", command=self.reset_list)
        self.menu_bar.add_command(label="关于", command=self.show_about)

        self.start_button = ttk.Button(
            self.root,
            text="开始",
            command=self.start_scrolling)
        self.start_button.place(x=100, y=400)

        self.save_button = ttk.Button(
            self.root,
            text="记录",
            command=self.save_data)
        self.save_button.place(x=220, y=400)

        self.revoke_button = ttk.Button(
            self.root,
            text="撤销",
            command=self.revoke_save)
        self.revoke_button.place(x=340, y=400)

        self.is_scrolling = False
        self.teacher_data = pd.DataFrame(columns=["工号", "姓名"])
        self.teacher_num = 0
        self.export_data = pd.DataFrame(columns=["序号", "工号", "姓名"])
        self.export_num = 0
        self.table_num = 0

        self.config_label = tk.Label(
            self.root,
            text=f"文件数: {self.table_num}\t人员数: {self.teacher_num}\n记录数: {self.export_num}\t剩余数: {self.teacher_num-self.export_num}",
            font=("Times New Roman", 12),
            justify=tk.LEFT)
        self.config_label.place(x=20, y=20)

        self.label = tk.Label(
            self.root,
            text="",
            font=("Times New Roman", 32),
            justify=tk.CENTER,
            width=17,
            height=2)
        self.label.place(x=20, y=150)

        self.result = tk.Text(
            self.root,
            height=22,
            width=24,
            font=("Times New Roman", 12))
        self.result.configure(state=tk.NORMAL)
        self.yscrollbar = Scrollbar(self.root)
        self.yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result.place(x=550, y=0)
        self.yscrollbar.config(command=self.result.yview)
        self.result.config(yscrollcommand=self.yscrollbar.set)
        self.result.insert(tk.END, "——— 已选记录 ———\n")
        self.result.configure(state=tk.DISABLED)

    def start_scrolling(self):
        """
        开始/停止滚动选择人员
        """
        if self.teacher_data.empty:
            messagebox.showwarning("错误", "请先[导入名单]")
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
        if len(teacher_id) < TEACHER_ID_LENGTH:
            for i in range(TEACHER_ID_LENGTH-len(teacher_id)):
                teacher_id = '0'+teacher_id
        self.label.config(text=f'{teacher_id} {teacher_name}')
        self.root.after(SCROLL_INTERVAL_TIME, self.scroll_names)

    def save_data(self):
        """
        记录已选人员信息
        """
        if self.is_scrolling:
            self.is_scrolling = False
            messagebox.showinfo("提示", "请[停止]后记录")
            self.is_scrolling = True
            self.scroll_names()
            return
        if not self.label["text"] or not self.teacher_num:
            messagebox.showwarning("错误", "未选择人员")
            return
        row = str(self.label["text"]).split()
        if len(row) < 2:
            messagebox.showwarning("错误", "数据不合法")
            return
        if row[0] in self.export_data.values:
            messagebox.showwarning("错误", "人员已[记录]")
            return
        self.export_data.loc[len(self.export_data)] = [self.export_num+1] + row
        self.export_num += 1
        self.result.configure(state=tk.NORMAL)
        self.result.insert(tk.END, f'{self.export_num} {row[0]} {row[1]}\n')
        self.result.focus_force()
        self.result.see(tk.END)
        self.result.configure(state=tk.DISABLED)
        self.update_config()

    def revoke_save(self):
        """
        撤销最近一次记录
        """
        if not self.export_num:
            messagebox.showwarning("错误", "无历史[记录]")
            return
        self.result.configure(state=tk.NORMAL)
        totalLen = len(self.result.get('1.0', tk.END).split("\n"))
        delstart = f"{totalLen-2}.0"
        delend = f"{totalLen}.0"
        self.result.delete(delstart, delend)
        self.export_data = self.export_data[:-1]
        self.export_num -= 1
        self.result.focus_force()
        self.result.insert(tk.END, '\n')
        self.result.configure(state=tk.DISABLED)
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
                self.teacher_data = pd.concat([self.teacher_data, data])
                self.teacher_num += len(data)
                self.table_num += 1
                self.update_config()
                messagebox.showinfo("上传成功",
                                    f"[导入名单]成功\n成功导入{self.teacher_num}条数据")
            except pd.errors.ParserError:
                messagebox.showwarning("错误", "无法解析文件\n请选择正确的Excel文件")

    def reset_list(self):
        """
        重置人员名单和记录信息
        """
        if not messagebox.askyesno("提示",
                                   "确定要[重置所有]吗?\n"
                                   "将清空: [导入名单], [导出名单], [已选记录]"):
            return
        self.teacher_data = pd.DataFrame(columns=["工号", "姓名"])
        self.export_data = pd.DataFrame(columns=["序号", "工号", "姓名"])
        self.teacher_num = 0
        self.export_num = 0
        self.table_num = 0
        self.label.config(text="")
        self.result.configure(state=tk.NORMAL)
        self.result.delete('1.0', tk.END)
        self.result.insert(tk.END, "——— 已选记录 ———\n")
        self.result.configure(state=tk.DISABLED)
        self.update_config()
        messagebox.showinfo("成功", "[重置所有]成功")

    def export_list(self):
        """
        导出已记录的人员信息
        """
        if self.export_data.empty:
            messagebox.showwarning("错误", "[导出名单]为空")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.export_data.to_excel(file_path, index=False)
            messagebox.showinfo("保存成功",
                                f"[导出名单]成功\n成功导出{self.export_num}条数据\n路径: {file_path}")

    def show_about(self):
        """
        显示关于窗口
        """
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        pad_width = round((self.screen_width-450)/2)
        pad_height = round((self.screen_height-250)/2)
        about_window.geometry(f"450x250+{pad_width}+{pad_height}")

        label = tk.Label(
            about_window,
            text="# 更新日志 2023.06.23\n"
                 "1. 优化按钮布局, 去除无用按钮, 增加撤销功能\n"
                 "2. 修复重置后记录功能失效\n"
                 "3. 支持多表导入, 添加人员数量显示\n"
                 "4. 重置和关闭窗口时增加提示, 避免误操作\n"
                 "\n"
                 "# 关于\n"
                 "- 作者:\tHaidong Xin\n"
                 "- Email:\txhd0728@hrbeu.edu.cn",
            font=("Times New Roman", 12),
            anchor=tk.W,
            justify=tk.LEFT)
        label.pack(pady=20, padx=20)

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
            text=f"文件数: {self.table_num}\t人员数: {self.teacher_num}\n记录数: {self.export_num}\t剩余数: {self.teacher_num-self.export_num}")

    def valid_data(self, df) -> bool:
        """
        验证数据是否合法
        """
        if '工号' in df.columns and '姓名' in df.columns:
            return True
        return False

    def window_close_handle(self):
        """
        窗口关闭监听函数
        """
        if messagebox.askyesno("提示",
                               "确认关闭软件吗?\n"
                               "将清空: [导入名单], [导出名单], [已选记录]"):
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.resizable(0, 0)
    root.mainloop()
