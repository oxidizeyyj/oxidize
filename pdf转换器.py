import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
import win32com.client

# 选择文件夹
def select_folder(folder_type):
    folder_selected = filedialog.askdirectory(title="请选择文件夹")
    if folder_selected:
        if folder_type == "input":
            folder_path.set(folder_selected)
            load_files_in_folder(folder_selected)  # 自动加载文件夹中的Word文档
        elif folder_type == "output":
            output_folder_path.set(folder_selected)

# 选择当前文件夹
def select_current_folder(folder_type):
    current_folder = os.path.dirname(os.path.abspath(__file__))  # 获取当前脚本所在的目录
    if folder_type == "input":
        folder_path.set(current_folder)
        load_files_in_folder(current_folder)  # 加载当前文件夹中的Word文件
    elif folder_type == "output":
        output_folder_path.set(current_folder)

# 自动加载文件夹中的 Word 文档
def load_files_in_folder(folder):
    # 列出该文件夹下的所有 .doc 和 .docx 文件
    files = [f for f in os.listdir(folder) if f.lower().endswith((".doc", ".docx"))]
    
    print(f"检测到以下文件：{files}")  # 打印出文件列表，用于调试

    files_listbox.delete(0, tk.END)  # 清空列表框
    if not files:
        messagebox.showwarning("警告", "该文件夹下没有找到 Word 文件！")
    for file in files:
        files_listbox.insert(tk.END, file)

# 开始替换并转换文件
def start_replace():
    input_folder = folder_path.get()
    output_folder = output_folder_path.get()

    if not input_folder or not os.path.exists(input_folder):
        messagebox.showerror("错误", "请选择有效的输入文件夹路径")
        return
    if not output_folder or not os.path.exists(output_folder):
        messagebox.showerror("错误", "请选择有效的输出文件夹路径")
        return

    # 获取勾选的文件
    selected_files = [files_listbox.get(i) for i in files_listbox.curselection()]
    if not selected_files:
        selected_files = [files_listbox.get(i) for i in range(files_listbox.size())]  # 如果没有选择，默认全选

    # 替换文档内容并转换为PDF
    total_files = len(selected_files)
    count = 0
    for filename in selected_files:
        full_input_path = os.path.join(input_folder, filename)
        try:
            doc = Document(full_input_path)
        except Exception as e:
            messagebox.showerror("错误", f"打开文件失败: {filename}\n{e}")
            continue

        # 转换为PDF
        pdf_path = os.path.splitext(os.path.join(output_folder, filename))[0] + ".pdf"
        used_app = docx_to_pdf(full_input_path, pdf_path)

        count += 1
        progress_percent = (count / total_files) * 100
        progress_var.set(progress_percent)
        progress_label.config(text=f"{count}/{total_files} (PDF)")

        root.update_idletasks()

    messagebox.showinfo("完成", f"文件处理完成，输出文件已保存至：\n{output_folder}")

# 使用WPS或Office将Word文档转换为PDF
def docx_to_pdf(input_path, output_path):
    try:
        try:
            wps = win32com.client.Dispatch('Kwps.Application')
            wps.Visible = False
            doc = wps.Documents.Open(input_path)
            doc.ExportAsFixedFormat(output_path, 17)
            doc.Close()
            wps.Quit()
            return "WPS"
        except Exception:
            pass

        try:
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)
            doc.Close()
            word.Quit()
            return "Office"
        except Exception:
            pass

        messagebox.showwarning(
            "未生成PDF",
            "未检测到 WPS 或 Microsoft Office，无法生成PDF文件"
        )
        return "无"
    except Exception as e:
        messagebox.showerror("错误", f"生成PDF时出错：\n{e}")
        return "错误"

# 选择所有文件
def select_all_files():
    files_listbox.select_set(0, tk.END)

# 取消选择所有文件
def deselect_all_files():
    files_listbox.select_clear(0, tk.END)

# 创建GUI界面
root = tk.Tk()
root.title("Word 批量转换为PDF")

folder_path = tk.StringVar()
output_folder_path = tk.StringVar()

# 输入文件夹
tk.Label(root, text="选择输入文件夹：").grid(row=0, column=0, sticky="w", padx=10, pady=5)
tk.Entry(root, textvariable=folder_path, width=50).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="浏览...", command=lambda: select_folder("input")).grid(row=0, column=2, padx=10, pady=5)
tk.Button(root, text="当前文件夹", command=lambda: select_current_folder("input")).grid(row=0, column=3, padx=10, pady=5)

# 输出文件夹
tk.Label(root, text="选择输出文件夹：").grid(row=1, column=0, sticky="w", padx=10, pady=5)
tk.Entry(root, textvariable=output_folder_path, width=50).grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="浏览...", command=lambda: select_folder("output")).grid(row=1, column=2, padx=10, pady=5)
tk.Button(root, text="当前文件夹", command=lambda: select_current_folder("output")).grid(row=1, column=3, padx=10, pady=5)

# 文件选择框
tk.Label(root, text="选择要处理的文件：").grid(row=2, column=0, sticky="w", padx=10, pady=5)
files_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=10)
files_listbox.grid(row=2, column=1, columnspan=2, padx=10, pady=5)

# 一键全选、取消全选
tk.Button(root, text="一键全选", command=select_all_files).grid(row=3, column=0, padx=10, pady=5)
tk.Button(root, text="取消全选", command=deselect_all_files).grid(row=3, column=1, padx=10, pady=5)

# 开始处理按钮
tk.Button(root, text="开始处理", command=start_replace).grid(row=4, column=1, pady=15)

# ---------- 添加进度条 ----------
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=400)
progress_bar.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

progress_label = tk.Label(root, text="0/0")
progress_label.grid(row=6, column=0, columnspan=3, padx=10, pady=5)
# ---------- 添加进度条结束 ----------

root.mainloop()
