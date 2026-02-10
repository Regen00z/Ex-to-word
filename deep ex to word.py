import tkinter as tk
from tkinter import filedialog, messagebox
import xlwings as xw
from docx import Document

# 选择 Excel 文件
def select_excel_file():
    file_path = filedialog.askopenfilename(
        title="选择 Excel 文件",
        filetypes=[("Excel 文件", "*.xlsx *.xls")]
    )
    if file_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, file_path)

# 选择 Word 模板文件
def select_word_template():
    file_path = filedialog.askopenfilename(
        title="选择 Word 模板文件",
        filetypes=[("Word 文件", "*.docx")]
    )
    if file_path:
        template_entry.delete(0, tk.END)
        template_entry.insert(0, file_path)

# 生成 Word 文件
def generate_word_files():
    excel_path = excel_entry.get()
    template_path = template_entry.get()
    output_folder = filedialog.askdirectory(title="选择输出文件夹")

    if not excel_path or not template_path or not output_folder:
        messagebox.showerror("错误", "请选择 Excel 文件、Word 模板文件和输出文件夹")
        return

    try:
        # 打开 Excel 文件
        app = xw.App(visible=False)
        workbook = app.books.open(excel_path)
        sheet = workbook.sheets[0]

        # 读取 Excel 数据
        data = sheet.used_range.value  # 获取所有数据
        headers = data[0]  # 第一行为表头
        rows = data[1:]  # 其余行为数据

        # 遍历每一行数据，生成 Word 文件
        for i, row in enumerate(rows):
            doc = Document(template_path)  # 打开 Word 模板
            for paragraph in doc.paragraphs:
                for header, value in zip(headers, row):
                    # 替换模板中的占位符，例如 {{Name}} 替换为实际值
                    placeholder = f"{{{{{header}}}}}"
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, str(value))

            # 保存生成的 Word 文件
            output_path = f"{output_folder}/output_{i + 1}.docx"
            doc.save(output_path)

        workbook.close()
        app.quit()
        messagebox.showinfo("成功", f"已生成 {len(rows)} 个 Word 文件到 {output_folder}")
    except Exception as e:
        messagebox.showerror("错误", f"生成 Word 文件时出错: {str(e)}")

# 创建 GUI
root = tk.Tk()
root.title("Excel 数据生成 Word 文件")

# Excel 文件选择
tk.Label(root, text="Excel 文件路径:").grid(row=0, column=0, padx=5, pady=5)
excel_entry = tk.Entry(root, width=50)
excel_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="选择文件", command=select_excel_file).grid(row=0, column=2, padx=5, pady=5)

# Word 模板选择
tk.Label(root, text="Word 模板路径:").grid(row=1, column=0, padx=5, pady=5)
template_entry = tk.Entry(root, width=50)
template_entry.grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="选择文件", command=select_word_template).grid(row=1, column=2, padx=5, pady=5)

# 生成按钮
tk.Button(root, text="生成 Word 文件", command=generate_word_files).grid(row=2, column=1, pady=10)

# 运行 GUI
root.mainloop()