import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import os


def word_to_pdf_with_gui():
    # 1. 初始化弹窗环境
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)  # 强制置顶，防止弹窗被挡住

    # 2. 选择输入的 Word 文件（支持 .doc 和 .docx）
    input_file = filedialog.askopenfilename(
        title="请选择要转换的 Word 文档",
        filetypes=[("Word 文档", "*.docx *.doc")]
    )

    if not input_file:
        return

    # 3. 提取原文件名，作为 PDF 的默认保存名
    default_filename = os.path.splitext(os.path.basename(input_file))[0] + ".pdf"

    # 4. 选择 PDF 保存位置
    output_file = filedialog.asksaveasfilename(
        title="请选择 PDF 保存位置",
        defaultextension=".pdf",
        initialfile=default_filename,
        filetypes=[("PDF 文件", "*.pdf")]
    )

    if not output_file:
        return

    # 5. 开始转换 (使用 win32com 引擎)
    # win32com 对路径非常严格，必须使用绝对路径
    abs_input = os.path.abspath(input_file)
    abs_output = os.path.abspath(output_file)

    word = None
    try:
        print(f"正在后台启动 Word 处理文件：{abs_input} ...")

        # 启动 Word 后台程序
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 保持静默，不显示 Word 界面

        # 打开文档
        doc = word.Documents.Open(abs_input)

        # 17 是 Word 内部代表 PDF 格式的代码 (wdFormatPDF)
        doc.SaveAs(abs_output, FileFormat=17)
        doc.Close()

        # 成功提示
        messagebox.showinfo("转换成功", f"文件已成功保存至：\n{abs_output}", parent=root)
        print("转换完成！")

    except Exception as e:
        # 错误提示
        error_msg = str(e)
        messagebox.showerror("转换失败", f"发生错误：\n{error_msg}", parent=root)
        print(f"转换失败: {error_msg}")

    finally:
        # ⚠️ 这一步极其重要：无论成功失败，必须彻底关闭 Word 进程
        if word:
            word.Quit()


if __name__ == "__main__":
    word_to_pdf_with_gui()