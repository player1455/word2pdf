import tkinter as tk
from converter.word2pdf import word_to_pdf_with_gui
from converter.pdf2word import pdf_to_word_with_gui

# 主程序入口
if __name__ == "__main__":
    # 创建主窗口
    root = tk.Tk()
    root.title("办公文档格式转换工具")

    # 设置窗口大小和居中显示
    window_width = 300
    window_height = 180
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

    # 强制窗口置顶，防止被其他软件挡住
    root.attributes('-topmost', True)

    # 添加界面文字说明
    tk.Label(root, text="请选择你要执行的操作", font=("微软雅黑", 12, "bold")).pack(pady=15)

    # 创建按钮容器，让按钮排版更整齐
    btn_frame = tk.Frame(root)
    btn_frame.pack()

    # 核心：选择按钮，点击后分别调用上面定义好的两个核心函数，并将 root 传递进去
    tk.Button(btn_frame, text="1. Word 转 PDF", font=("微软雅黑", 10), width=18,
              command=lambda: word_to_pdf_with_gui(root)).pack(pady=8)

    tk.Button(btn_frame, text="2. PDF 转 Word", font=("微软雅黑", 10), width=18,
              command=lambda: pdf_to_word_with_gui(root)).pack(pady=8)

    # 启动窗口主循环
    root.mainloop()