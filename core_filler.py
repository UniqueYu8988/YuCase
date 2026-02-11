import tkinter as tk
from tkinter import ttk, messagebox
import json
import re
import os
import time
import threading
import ctypes
from ctypes import wintypes
import sys

# --- Windows API Definitions for Hotkey & Input ---
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

VK_F9 = 0x78
VK_TAB = 0x09
VK_CONTROL = 0x11
VK_V = 0x56
KEYEVENTF_KEYUP = 0x0002

class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD), 
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD), 
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG))]

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD), ("wParamH", wintypes.WORD)]

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]
    _anonymous_ = ("_input",)
    _fields_ = [("type", wintypes.DWORD), ("_input", _INPUT)]

def press_key(hexKeyCode):
    x = INPUT(type=1, ki=KEYBDINPUT(wVk=hexKeyCode))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

def release_key(hexKeyCode):
    x = INPUT(type=1, ki=KEYBDINPUT(wVk=hexKeyCode, dwFlags=KEYEVENTF_KEYUP))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

def send_paste_and_tab():
    press_key(VK_CONTROL)
    press_key(VK_V)
    release_key(VK_V)
    release_key(VK_CONTROL)
    time.sleep(0.05)
    press_key(VK_TAB)
    release_key(VK_TAB)

# --- Field Metadata Definition ---
FIELD_ORDER = [
    # Module 1
    ("field_1_3", "病案号"),
    ("field_1_1", "医疗付款方式"),
    ("ph_auto", "占位"),
    ("field_1_2", "第 N 次住院"),
    ("field_1_3", "病案号"),
    
    ("field_1_4", "姓名"),
    ("field_1_5", "性别"),
    ("field_1_6", "出生日期"),
    ("field_1_7", "年龄"),
    ("field_1_8", "国籍"),
    
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("field_1_9_native", "籍贯"),

    ("field_1_10", "民族"),
    ("field_1_11", "身份证号"),
    ("field_1_12", "职业"),
    ("field_1_13", "婚姻"),
    ("field_1_9", "出生地"),
    
    ("field_1_15", "电话"),
    ("ph_auto", "占位"),
    ("field_1_14", "出生地(2)"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("field_1_17", "联系人姓名"),
    ("field_1_18", "关系"),
    ("field_1_16", "出生地(3)"),
    
    ("field_1_15", "电话"),
    ("field_1_21", "入院途径"),
    ("ph_auto", "占位"),
    ("field_1_22_1", "入院时间"),
    ("field_1_22_2", "入院时"),
    
    ("field_1_23", "入院科别"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("field_1_24_1", "出院时间"),
    ("field_1_24_2", "出院时"),
    
    ("field_1_25", "出院科别"),
    ("ph_auto", "占位"),
    ("field_1_26", "实际住院（天）"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    
    ("field_1_27", "门(急)诊诊断"),
    ("field_1_28", "疾病编码"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    
    ("field_1_29", "入院病情"),
    ("field_1_27", "门(急)诊诊断"),
    ("field_1_28", "疾病编码"),
    ("field_1_29", "入院病情"),

    # Module 2
    ("field_2_1", "科主任"),
    ("field_2_2", "主(副)任医师"),
    ("field_2_3", "主治医师"),
    ("field_2_4", "住院医师"),
    
    ("field_2_5", "责任护士"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),

    ("field_2_6", "病案质量"),
    ("field_2_7", "质控医师"),
    ("field_2_8", "质控护士"),
    ("field_2_9", "质控日期"),
    ("field_2_10", "病案质量(代码)"),

    # Module 3
    ("field_3_1", "总费用"),
    ("field_3_2", "自付金额"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    ("ph_auto", "占位"),
    
    ("field_3_3", "一般治疗费"),
    ("field_3_4", "护理费"),
    ("field_3_5", "其他费用"),
    ("ph_auto", "占位"),
    ("field_3_6", "实验室诊断"),
    ("field_3_7", "影像学诊断"),
    
    ("field_3_8", "西药费"),
    ("ph_auto", "占位"),
    ("field_3_9", "中成药费"),
    ("ph_auto", "占位"),
    ("field_3_10", "中草药费"),
    ("field_3_11", "其他费")
]

# --- UI Utilities ---
def round_rectangle(x1, y1, x2, y2, radius=25, **kwargs):
    points = [x1+radius, y1,
              x1+radius, y1,
              x2-radius, y1,
              x2-radius, y1,
              x2, y1,
              x2, y1+radius,
              x2, y1+radius,
              x2, y2-radius,
              x2, y2-radius,
              x2, y2,
              x2-radius, y2,
              x2-radius, y2,
              x1+radius, y2,
              x1+radius, y2,
              x1, y2,
              x1, y2-radius,
              x1, y2-radius,
              x1, y1+radius,
              x1, y1+radius,
              x1, y1]
    return points

class RoundedFrame(tk.Canvas):
    def __init__(self, parent, width=400, height=200, radius=20, color="#FFFFFF", border_color="#E5E7EB", **kwargs):
        tk.Canvas.__init__(self, parent, width=width, height=height, borderwidth=0, 
                           relief="flat", highlightthickness=0, bg=parent["bg"], **kwargs)
        self.radius = radius
        self.color = color
        
        # Draw rounded background
        self.create_polygon(round_rectangle(2, 2, width-2, height-2, radius=radius), 
                            fill=color, outline=border_color, smooth=True, width=2)
        
        # Inner frame to hold widgets
        self.inner = tk.Frame(self, bg=color)
        self.create_window(width/2, height/2, window=self.inner, width=width-radius, height=height-radius)

    def pack(self, **kwargs):
        tk.Canvas.pack(self, **kwargs)

class RoundedLabel(tk.Canvas):
    def __init__(self, parent, width=300, height=60, radius=20, bg_color="#FFF7ED", fg_color="#000000", 
                 text="", font=("Microsoft YaHei UI", 16)):
        tk.Canvas.__init__(self, parent, width=width, height=height, borderwidth=0, 
                           relief="flat", highlightthickness=0, bg=parent["bg"])
        self.radius = radius
        self.bg_color = bg_color
        
        # Draw rounded rect
        self.rect = self.create_polygon(round_rectangle(2, 2, width-2, height-2, radius=radius), 
                                        fill=bg_color, outline="#FDBA74", smooth=True, width=1)
        
        self.text_id = self.create_text(20, height/2, text=text, fill=fg_color, font=font, anchor="w")

    def config(self, text=None):
        if text is not None:
            self.itemconfig(self.text_id, text=text)

class RoundedButton(tk.Canvas):
    def __init__(self, parent, width=120, height=55, corner_radius=25, 
                 padding=0, color="#22c55e", text="Button", text_color="white", 
                 font=("Microsoft YaHei UI", 14, "bold"), command=None):
        tk.Canvas.__init__(self, parent, borderwidth=0, 
                           relief="flat", highlightthickness=0, bg=parent["bg"])
        self.command = command
        self.color = color
        self.hover_color = self.adjust_brightness(color, 0.9) # Slightly darker
        self.text_color = text_color
        self.width = width
        self.height = height
        self.corner_radius = corner_radius
        self.font = font
        self.text = text

        self.configure(width=width, height=height)
        
        # Draw initial state
        self.rect_id = self.create_polygon(
            round_rectangle(2, 2, width-2, height-2, radius=corner_radius), 
            fill=self.color, outline="", smooth=True, width=0)
        
        self.text_id = self.create_text(width/2, height/2, text=text, fill=text_color, font=font)

        # Bind events
        self.bind("<Button-1>", self.on_click)
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def adjust_brightness(self, hex_color, factor):
        # Simple brightness adjustment
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        r = max(0, min(255, int(r * factor)))
        g = max(0, min(255, int(g * factor)))
        b = max(0, min(255, int(b * factor)))
        return f"#{r:02x}{g:02x}{b:02x}"

    def on_click(self, event):
        if self.command:
            self.command()

    def on_enter(self, event):
        self.itemconfig(self.rect_id, fill=self.hover_color)

    def on_leave(self, event):
        self.itemconfig(self.rect_id, fill=self.color)
        
    def configure_color(self, color):
        self.color = color
        self.hover_color = self.adjust_brightness(color, 0.9)
        self.itemconfig(self.rect_id, fill=self.color)

    def configure_text(self, text):
        self.itemconfig(self.text_id, text=text)

class AutoFillerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("YuCase 病历助手")
        self.root.geometry("540x800")
        self.root.configure(bg="#FFF7ED")
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass

        self.base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.getcwd()
        
        self.records = []
        self.current_record = {}
        self.current_field_index = 0
        self.listening = False
        self.listener_thread = None
        
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        # --- Header ---
        header_frame = tk.Frame(self.root, bg="#FFF7ED", pady=25)
        header_frame.pack(fill="x")
        
        tk.Label(header_frame, text="YuCase 自动化填表", 
                 font=("Microsoft YaHei UI", 24, "bold"), 
                 bg="#FFF7ED", fg="#431407").pack()
        
        tk.Label(header_frame, text="按 F9 一键自动填写", 
                 font=("Microsoft YaHei UI", 12), 
                 bg="#FFF7ED", fg="#9A3412").pack(pady=(5, 0))

        # --- Section 1: Select ---
        content_frame = tk.Frame(self.root, bg="#FFF7ED", padx=35)
        content_frame.pack(fill="both", expand=True)

        tk.Label(content_frame, text="1. 请选择一份病历：", font=("Microsoft YaHei UI", 14, "bold"), 
                 bg="#FFF7ED", fg="#7C2D12", anchor="w").pack(fill="x", pady=(20, 10))
        
        # Custom Combobox Style - Large
        style = ttk.Style()
        style.theme_use('vista') 
        style.configure("TCombobox", fieldbackground="#FFEDD5", background="#FFF7ED", arrowsize=30)
        self.root.option_add("*TCombobox*Listbox*Font", ("Microsoft YaHei UI", 14))
        
        self.combo_records = ttk.Combobox(content_frame, state="readonly", font=("Microsoft YaHei UI", 14), height=12)
        self.combo_records.pack(fill="x", ipady=10)
        self.combo_records.bind("<<ComboboxSelected>>", self.on_record_select)

        # --- Section 2: Info Card ---
        
        card_inner = tk.Frame(content_frame, bg="#FFFAF0", padx=20, pady=20)
        # Using a simple frame for container 
        card_inner_container = tk.Frame(content_frame, bg="#FDBA74", padx=2, pady=2)
        card_inner_container.pack(fill="x", pady=25)
        card_inner = tk.Frame(card_inner_container, bg="#FFFAF0", padx=20, pady=20)
        card_inner.pack(fill="x")
        
        # Current Field & Value
        tk.Label(card_inner, text="当前准备填入：", font=("Microsoft YaHei UI", 12), bg="#FFFAF0", fg="#9A3412").pack(anchor="w")
        self.lbl_current_field = tk.Label(card_inner, text="--", font=("Microsoft YaHei UI", 18, "bold"), bg="#FFFAF0", fg="#431407")
        self.lbl_current_field.pack(anchor="w", pady=(5, 15))
        
        tk.Label(card_inner, text="内容预览：", font=("Microsoft YaHei UI", 12), bg="#FFFAF0", fg="#9A3412").pack(anchor="w")
        
        # Rounded Label for Value
        self.lbl_current_value = RoundedLabel(card_inner, width=350, height=80, radius=20, bg_color="#FFF7ED", text="--")
        self.lbl_current_value.pack(pady=(5, 0))

        # Reset Button
        self.btn_reset = tk.Button(card_inner, text="↺ 重新开始", font=("Microsoft YaHei UI", 11), 
                                   bg="#FFFAF0", fg="#EA580C", bd=0, cursor="hand2", activebackground="#FFFAF0", 
                                   activeforeground="#C2410C", command=self.reset_index)
        self.btn_reset.pack(anchor="e", pady=(10,0))

        # --- Section 3: Actions ---
        action_frame = tk.Frame(self.root, bg="#FFF7ED")
        action_frame.pack(fill="x", padx=35, pady=(0, 40))

        # We'll use grid for buttons to place them nicely
        action_frame.columnconfigure(0, weight=3) # Larger Button
        action_frame.columnconfigure(1, weight=2)

        # Toggle Button (Green -> Warm Green)
        self.btn_toggle = RoundedButton(action_frame, text="启动监听 (F9)", color="#16A34A", width=220, height=60, command=self.toggle_listening)
        self.btn_toggle.grid(row=0, column=0, padx=(0, 15), sticky="ew")

        # Save Button (Orange)
        self.btn_save = RoundedButton(action_frame, text="保存数据", color="#EA580C", width=140, height=60, command=self.save_data_from_clipboard)
        self.btn_save.grid(row=0, column=1, padx=(15, 0), sticky="ew")
        
        self.lbl_status = tk.Label(self.root, text="就绪", font=("Microsoft YaHei UI", 12), bg="#FFF7ED", fg="#9A3412")
        self.lbl_status.pack(side="bottom", pady=15)

    def load_data(self):
        html_path = os.path.join(self.base_path, 'medical_record_lite.html')
        if not os.path.exists(html_path):
            messagebox.showerror("错误", "找不到 medical_record_lite.html 文件！")
            return

        try:
            with open(html_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            match = re.search(r'let\s+records\s*=\s*(\[.*?\])\s*;', content, re.DOTALL)
            if match:
                json_str = match.group(1)
                json_str = re.sub(r'//.*', '', json_str) 
                self.records = json.loads(json_str)
                
                items = []
                for idx, r in enumerate(self.records):
                    name = r.get("field_1_4", "未知姓名")
                    items.append(f"{idx+1:02d}. {name}")
                self.combo_records['values'] = items
                if items:
                    self.combo_records.current(0)
                    self.on_record_select(None)
            else:
                self.lbl_status.config(text="错误：无法读取 HTML 数据")

        except Exception as e:
            messagebox.showerror("异常", f"读取数据失败: {str(e)}")

    def on_record_select(self, event):
        idx = self.combo_records.current()
        if idx >= 0:
            self.current_record = self.records[idx]
            self.reset_index()

    def reset_index(self):
        self.current_field_index = 0
        self.update_field_display()

    def update_field_display(self):
        if self.current_field_index < len(FIELD_ORDER):
            key, name = FIELD_ORDER[self.current_field_index]
            val = self.current_record.get(key, "")
            self.lbl_current_field.config(text=f"{self.current_field_index + 1}. {name}")
            self.lbl_current_value.config(text=val)
        else:
            self.lbl_current_field.config(text="完成")
            self.lbl_current_value.config(text="所有字段已填完")

    def toggle_listening(self):
        if not self.listening:
            self.listening = True
            self.btn_toggle.configure_text("停止监听")
            self.btn_toggle.configure_color("#ef4444")
            self.lbl_status.config(text=">>> 正在监听 F9 键... <<<", fg="#22c55e")
            self.listener_thread = threading.Thread(target=self.listen_loop, daemon=True)
            self.listener_thread.start()
        else:
            self.listening = False
            self.btn_toggle.configure_text("启动监听 (F9)")
            self.btn_toggle.configure_color("#22c55e")
            self.lbl_status.config(text="已暂停", fg="#94a3b8")

    def listen_loop(self):
        was_pressed = False
        while self.listening:
            state = user32.GetAsyncKeyState(VK_F9)
            is_down = (state & 0x8000) != 0
            
            if is_down and not was_pressed:
                self.perform_fill_action()
                was_pressed = True
            elif not is_down:
                was_pressed = False
            
            time.sleep(0.05)

    def perform_fill_action(self):
        if self.current_field_index >= len(FIELD_ORDER):
            return

        key, name = FIELD_ORDER[self.current_field_index]
        val = self.current_record.get(key, "")
        
        self.root.clipboard_clear()
        self.root.clipboard_append(val)
        self.root.update()
        
        send_paste_and_tab()
        
        self.current_field_index += 1
        self.root.after(0, self.update_field_display)

    def save_data_from_clipboard(self):
        try:
            data_str = self.root.clipboard_get()
            new_records = json.loads(data_str)
            if not isinstance(new_records, list):
                raise ValueError
        except:
             messagebox.showwarning("无数据", "请先在网页点击【保存修改】复制数据到剪贴板！")
             return

        if not messagebox.askyesno("确认", f"确定要更新 {len(new_records)} 条病历数据吗？"):
            return

        html_path = os.path.join(self.base_path, 'medical_record_lite.html') # PORTABLE CHANGE
        try:
            with open(html_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            pattern = re.compile(r'(let\s+records\s*=\s*)(\[.*?\])(\s*;)', re.DOTALL)
            new_json_str = json.dumps(new_records, ensure_ascii=False, indent=4)
            
            if pattern.search(content):
                new_content = pattern.sub(r'\1' + new_json_str.replace('\\', '\\\\') + r'\3', content)
                with open(html_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                messagebox.showinfo("成功", "保存成功！请刷新网页。")
                self.load_data()
            else:
                 messagebox.showerror("错误", "无法定位文件结构")

        except Exception as e:
            messagebox.showerror("失败", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoFillerApp(root)
    root.mainloop()
