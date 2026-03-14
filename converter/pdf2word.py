from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os

# 引入美化 Word 需要的库
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn


def beautify_generated_word(docx_path):
    """专门用于清洗和美化转换出来的 Word 文件"""
    print("正在清理和美化生成的 Word 文档...")
    doc = Document(docx_path)

    for para in doc.paragraphs:
        # 1. 去除段落首尾的无用空格
        para.text = para.text.strip()

        # 2. 统一行距为 1.5 倍
        para.paragraph_format.line_spacing = 1.5

        # 3. 统一字体格式
        for run in para.runs:
            # 英文字体设为 Arial
            run.font.name = 'Arial'
            # 强行设置中文字体为 微软雅黑 (Word底层严格要求这样设置中文字体)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

    doc.save(docx_path)


def pdf_to_word_with_gui(root):
    root.withdraw()

    input_file = filedialog.askopenfilename(
        title="请选择要转换的 PDF 文档",
        filetypes=[("PDF 文件", "*.pdf")]
    )
    if not input_file:
        root.deiconify()
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
        # 第一步：基础转换
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()

        # 第二步：调用我们写好的美化清洗函数
        beautify_generated_word(output_file)

        messagebox.showinfo("转换成功", f"文件已转换并美化，保存至：\n{output_file}")
    except Exception as e:
        messagebox.showerror("转换失败", f"发生错误：\n{str(e)}")
    finally:
        root.deiconify()