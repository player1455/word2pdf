from tkinter import filedialog, messagebox
import win32com.client
import os


def word_to_pdf_with_gui(root):
    root.withdraw()  # 隐藏主菜单

    input_file = filedialog.askopenfilename(
        title="请选择要转换的 Word 文档",
        filetypes=[("Word 文档", "*.docx *.doc")]
    )
    if not input_file:
        root.deiconify()  # 用户取消，恢复主菜单
        return

    default_filename = os.path.splitext(os.path.basename(input_file))[0] + ".pdf"
    output_file = filedialog.asksaveasfilename(
        title="请选择 PDF 保存位置",
        defaultextension=".pdf",
        initialfile=default_filename,
        filetypes=[("PDF 文件", "*.pdf")]
    )
    if not output_file:
        root.deiconify()
        return

    abs_input = os.path.abspath(input_file)
    abs_output = os.path.abspath(output_file)

    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(abs_input)
        doc.SaveAs(abs_output, FileFormat=17)
        doc.Close()
        messagebox.showinfo("转换成功", f"文件已成功保存至：\n{abs_output}")
    except Exception as e:
        messagebox.showerror("转换失败", f"发生错误：\n{str(e)}")
    finally:
        if word:
            word.Quit()
        root.deiconify()  # 恢复主菜单