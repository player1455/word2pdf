from tkinter import filedialog, messagebox
import win32com.client
import os


def word_to_pdf_with_gui(root):
    root.withdraw()

    input_file = filedialog.askopenfilename(
        title="请选择要转换的 Word 文档",
        filetypes=[("Word 文档", "*.docx *.doc")]
    )
    if not input_file:
        root.deiconify()
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

        # =================【美化排版逻辑】=================
        print("正在进行 PDF 排版美化...")
        # 1. 设置全文行距为 1.5 倍 (内部代码 5 代表 1.5 倍行距)
        doc.Content.ParagraphFormat.LineSpacingRule = 5
        # 2. 设置全文段落为两端对齐 (内部代码 3 代表两端对齐)
        doc.Content.ParagraphFormat.Alignment = 3
        # 3. 统一英文字体为 Arial，中文字体为 微软雅黑
        doc.Content.Font.NameAscii = "Arial"
        doc.Content.Font.NameFarEast = "微软雅黑"
        # ==================================================

        # 导出为 PDF
        doc.SaveAs(abs_output, FileFormat=17)

        # 关闭文档，【切记不要保存对原 Word 的修改】，0 代表 wdDoNotSaveChanges
        doc.Close(SaveChanges=0)

        messagebox.showinfo("转换成功", f"文件已美化并保存至：\n{abs_output}")
    except Exception as e:
        messagebox.showerror("转换失败", f"发生错误：\n{str(e)}")
    finally:
        if word:
            word.Quit()
        root.deiconify()