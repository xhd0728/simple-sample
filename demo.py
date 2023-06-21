import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

import pandas as pd
import ctypes


class App:
    def __init__(self, root):
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

        self.menu_bar.add_command(label="关于", command=self.show_about)

        self.start_button = ttk.Button(self.root,
                                       text="开始",
                                       command=self.start_scrolling)
        self.start_button.place(x=200, y=350)

        self.stop_button = ttk.Button(self.root,
                                      text="停止",
                                      command=self.stop_scrolling)
        self.stop_button.place(x=300, y=350)

        self.save_button = ttk.Button(self.root,
                                      text="记录",
                                      command=self.save_data)
        self.save_button.place(x=400, y=350)

        self.reset_button = ttk.Button(self.root,
                                       text="复位",
                                       command=self.reset_list)
        self.reset_button.place(x=500, y=350)

        self.is_scrolling = False
        self.teacher_data = pd.DataFrame(columns=["工号", "姓名"])
        self.teacher_num = 0
        self.export_data = pd.DataFrame(columns=["工号", "姓名"])
        self.export_num = 0

        self.config_lable = tk.Label(self.root,
                                     text=f"人员总数: {self.teacher_num}, 已记录: {self.export_num}",
                                     font=("Times New Roman", 16),
                                     justify="center")
        self.config_lable.place(x=20, y=20)

        self.label = tk.Label(self.root,
                              text="",
                              font=("Times New Roman", 42),
                              justify="center")
        self.label.place(x=100, y=150)

        self.result = tk.Text(self.root,
                              height=4,
                              width=40,
                              font=("Times New Roman", 12))
        self.result.place(x=200, y=400)

    def start_scrolling(self):
        if self.teacher_data.empty:
            messagebox.showwarning("错误", "请先导入人员名单")
            return
        if self.is_scrolling:
            return
        self.is_scrolling = True
        self.scroll_names()

    def stop_scrolling(self):
        self.is_scrolling = False

    def scroll_names(self):
        if not self.is_scrolling:
            return
        row = self.get_random_row()
        teacher_id = str(row.iat[0, 1])
        teacher_name = str(row.iat[0, 2])
        if len(teacher_id) == 9:
            teacher_id = '0'+teacher_id
        self.label.config(text=f'{teacher_id} {teacher_name}')
        self.root.after(100, self.scroll_names)

    def save_data(self):
        if self.is_scrolling:
            self.is_scrolling = False
        if not self.label["text"]:
            messagebox.showwarning("错误", "未选择人员")
            return
        row = self.label["text"].split()
        self.export_data.loc[len(self.export_data)] = row
        self.result.insert('end', f'{row[0]} {row[1]}\n')
        self.export_num += 1
        self.update_config()

    def import_list(self):
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
        self.teacher_data = pd.DataFrame()
        self.export_data = pd.DataFrame()
        self.teacher_num = 0
        self.export_num = 0
        self.label.config(text="")
        self.result.delete('1.0', 'end')
        self.update_config()
        messagebox.showinfo("成功", "重置成功")

    def export_list(self):
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
        return self.teacher_data.sample(n=1)

    def update_config(self):
        self.config_lable.config(
            text=f"人员总数: {self.teacher_num}, 已记录: {self.export_num}")

    def valid_data(self, df) -> bool:
        if '工号' in df.columns and '姓名' in df.columns:
            return True
        return False


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.resizable(0, 0)
    root.mainloop()
