import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import threading
import subprocess
import core_import
import core_filler
import ctypes

# 设置高 DPI 适配
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

class YuCaseLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("YuCase 病历管理集成系统")
        self.root.geometry("450x550")
        self.root.configure(bg="#FFF7ED")
        self.root.resizable(False, False)

        self.setup_ui()

    def setup_ui(self):
        # 头部标题
        header_frame = tk.Frame(self.root, bg="#FFF7ED", pady=30)
        header_frame.pack(fill="x")

        tk.Label(header_frame, text="YuCase 智能集成版", 
                 font=("Microsoft YaHei UI", 22, "bold"), 
                 bg="#FFF7ED", fg="#431407").pack()
        
        tk.Label(header_frame, text="请按照步骤点击下方按钮", 
                 font=("Microsoft YaHei UI", 11), 
                 bg="#FFF7ED", fg="#9A3412").pack(pady=(5, 0))

        # 按钮容器
        btn_frame = tk.Frame(self.root, bg="#FFF7ED", padx=50)
        btn_frame.pack(fill="both", expand=True)

        # 按钮 1：更新数据
        self.btn_update = core_filler.RoundedButton(
            btn_frame, text="1. 更新病历数据", color="#F97316", 
            width=350, height=70, command=self.run_update
        )
        self.btn_update.pack(pady=15)
        tk.Label(btn_frame, text="扫描文件夹中的 Word 并更新到网页", 
                 font=("Microsoft YaHei UI", 9), bg="#FFF7ED", fg="#C2410C").pack()

        # 按钮 2：查看网页
        self.btn_view = core_filler.RoundedButton(
            btn_frame, text="2. 查看病历列表", color="#8B5CF6", 
            width=350, height=70, command=self.open_html
        )
        self.btn_view.pack(pady=15)
        tk.Label(btn_frame, text="在浏览器中查看和修改所有解析出的数据", 
                 font=("Microsoft YaHei UI", 9), bg="#FFF7ED", fg="#6D28D9").pack()

        # 按钮 3：启动助手
        self.btn_assist = core_filler.RoundedButton(
            btn_frame, text="3. 启动填表助手", color="#16A34A", 
            width=350, height=70, command=self.start_filler
        )
        self.btn_assist.pack(pady=15)
        tk.Label(btn_frame, text="启动自动填表窗口，按 F9 自动填写", 
                 font=("Microsoft YaHei UI", 9), bg="#FFF7ED", fg="#15803D").pack()

        # 版权信息
        tk.Label(self.root, text="Designed for Mom | YuCase v1.0", 
                 font=("Microsoft YaHei UI", 9), bg="#FFF7ED", fg="#94a3b8").pack(side="bottom", pady=20)

    def get_base_path(self):
        # 即使被打包成 EXE，也能正确找到旁边的文件
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.getcwd()

    def run_update(self):
        # 在新线程中运行，防止界面卡死
        def task():
            try:
                self.btn_update.configure_text("正在更新...")
                # 确保 core_import 也在正确的目录下运行
                original_cwd = os.getcwd()
                os.chdir(self.get_base_path())
                core_import.main()
                os.chdir(original_cwd)
                messagebox.showinfo("成功", "病历数据更新完成！")
            except Exception as e:
                messagebox.showerror("错误", f"更新失败: {e}")
            finally:
                self.btn_update.configure_text("1. 更新病历数据")
        
        threading.Thread(target=task, daemon=True).start()

    def open_html(self):
        html_path = os.path.join(self.get_base_path(), "medical_record_lite.html")
        if os.path.exists(html_path):
            os.startfile(html_path)
        else:
            messagebox.showerror("错误", f"找不到文件：\n{html_path}\n请确保文件在软件同一目录下。")

    def start_filler(self):
        try:
            # 修复：打包后 sys.executable 是 EXE 本身，直接跑会重开
            # 我们通过传递参数 --filler 让 EXE 知道这次是要启动助手
            subprocess.Popen([sys.executable, "--filler"])
        except Exception as e:
            messagebox.showerror("错误", f"启动助手失败: {e}")

if __name__ == "__main__":
    # 如果检测到 --filler 参数，直接启动助手，不显示主界面
    if len(sys.argv) > 1 and sys.argv[1] == "--filler":
        root = tk.Tk()
        app = core_filler.AutoFillerApp(root)
        root.mainloop()
    else:
        root = tk.Tk()
        app = YuCaseLauncher(root)
        root.mainloop()
