import tkinter as tk
from tkinter import filedialog, messagebox
import os
from text import main
import sys

def resource_path(relative_path):
    # 兼容打包和开发环境
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# 假设码数表和格式模版在当前目录下
DEFAULT_SIZE_PATH = resource_path("码数表.xlsx")
DEFAULT_TEMPLATE_PATH = resource_path("格式模版.xlsx")

def run_main(student_file_path, output_dir):
    try:
        main(student_file_path, DEFAULT_SIZE_PATH, DEFAULT_TEMPLATE_PATH, output_dir)
        messagebox.showinfo("成功", f"处理完成，结果已保存到：{output_dir}")
    except Exception as e:
        messagebox.showerror("错误", str(e))

def select_student_file():
    file_path = filedialog.askopenfilename(title="选择身高体重表", filetypes=[("Excel文件", "*.xlsx")])
    student_entry.delete(0, tk.END)
    student_entry.insert(0, file_path)

def select_output_dir():
    dir_path = filedialog.askdirectory(title="选择输出文件夹")
    output_entry.delete(0, tk.END)
    output_entry.insert(0, dir_path)

def start_process():
    student_file = student_entry.get()
    output_dir = output_entry.get()
    if not student_file or not output_dir:
        messagebox.showwarning("提示", "请先选择学生表和输出路径！")
        return
    run_main(student_file, output_dir)

root = tk.Tk()
root.title("学校配码批量处理")

tk.Label(root, text="选择身高体重表:").grid(row=0, column=0, padx=10, pady=10)
student_entry = tk.Entry(root, width=40)
student_entry.grid(row=0, column=1, padx=10)
tk.Button(root, text="浏览", command=select_student_file).grid(row=0, column=2, padx=10)

tk.Label(root, text="选择输出文件夹:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=40)
output_entry.grid(row=1, column=1, padx=10)
tk.Button(root, text="浏览", command=select_output_dir).grid(row=1, column=2, padx=10)

tk.Button(root, text="开始处理", command=start_process, width=20).grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()