from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os


def pdf_to_word_with_gui(root):
    root.withdraw()  # 隐藏主菜单

    input_file = filedialog.askopenfilename(
        title="请选择要转换的 PDF 文档",
        filetypes=[("PDF 文件", "*.pdf")]
    )
    if not input_file:
        root.deiconify()  # 用户取消，恢复主菜单
        return

    default_filename = os.path.splitext(os.path.basename(input_file))[0] + ".docx"
    output_file = filedialog.asksaveasfilename(
        title="请选择 Word 保存位置",
        defaultextension=".docx",
        initialfile=default_filename,
        filetypes=[("Word 文档", "*.docx")]
    )
    if not output_file:
        root.deiconify()
        return

    try:
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()
        messagebox.showinfo("转换成功", f"文件已成功保存至：\n{output_file}")
    except Exception as e:
        messagebox.showerror("转换失败", f"发生错误：\n{str(e)}")
    finally:
        root.deiconify()  # 恢复主菜单