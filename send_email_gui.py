import os
import re
import json
import time
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
import threading

try:
    import customtkinter as ctk
    USE_CUSTOMTK = True
except ImportError:
    USE_CUSTOMTK = False


def get_color(color_name):
    """根据主题返回对比度合适的颜色"""
    if not USE_CUSTOMTK:
        return color_name
    # customtkinter 浅色主题下使用深色以确保对比度
    color_map = {
        'lightblue': '#4a90d9',
        'green': '#2d8a2d',
        'orange': '#cc7000',
        'red': '#c93030',
        'lightgreen': '#3cb043',
        'blue': '#4a90d9',
        'gray': '#6c757d'
    }
    return color_map.get(color_name, color_name)


FIELD_TYPES = {
    'key': '关键字段（匹配文件名编号）',
    'name': '姓名/称呼字段',
    'email': '邮件地址字段',
    'ignore': '不使用此字段'
}

FIELD_TYPE_KEYS = list(FIELD_TYPES.keys())


def extract_key_from_filename(filename):
    name_without_ext = os.path.splitext(filename)[0]
    match = re.match(r'^(\d+)', name_without_ext)
    if match:
        return match.group(1)
    return None


def suggest_field_type(header):
    if not header:
        return 'ignore'
    h = str(header).lower()
    if '学号' in h or '编号' in h or '工号' in h or 'id' in h:
        return 'key'
    elif '姓名' in h or '名字' in h:
        return 'name'
    elif '邮箱' in h or '邮件' in h or 'e-mail' in h:
        return 'email'
    elif '电话' in h or '手机' in h:
        return 'ignore'
    return 'ignore'


class ColumnMapping:
    def __init__(self):
        self.columns = []
        self.mapping = {}

    def load_from_excel(self, file_path):
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        self.columns = []
        self.mapping = {}

        for i, cell in enumerate(ws[1], 1):
            col_letter = openpyxl.utils.get_column_letter(i)
            header_value = cell.value
            if isinstance(header_value, (int, float)):
                header_value = str(int(header_value))
            self.columns.append({
                'letter': col_letter,
                'name': header_value or f'列{i}',
                'suggested': suggest_field_type(header_value)
            })
        wb.close()

    def set_mapping(self, letter, field_type):
        for existing_letter, existing_type in list(self.mapping.items()):
            if existing_type == 'key' and field_type != 'key' and letter != existing_letter:
                pass
            if existing_type == field_type and letter != existing_letter:
                if messagebox.askyesno("确认", f"'{FIELD_TYPES[existing_type]}' 已分配给其他列，是否更换？"):
                    del self.mapping[existing_letter]
                    break
                else:
                    return False
        self.mapping[letter] = field_type
        return True

    def get_key_column(self):
        for letter, type_ in self.mapping.items():
            if type_ == 'key':
                return letter
        return None

    def get_column_index(self, letter):
        if not letter:
            return -1
        col_idx = openpyxl.utils.column_index_from_string(letter)
        return col_idx - 1

    def get_data_for_field(self, row_data, field_type):
        if field_type not in self.mapping:
            return ''
        letter = None
        for ltr, typ in self.mapping.items():
            if typ == field_type:
                letter = ltr
                break
        if not letter:
            return ''
        col_idx = self.get_column_index(letter)
        if col_idx < 0 or col_idx >= len(row_data):
            return ''
        value = row_data[col_idx]
        if isinstance(value, (int, float)):
            value = str(int(value))
        return value.strip() if value else ''


class RecipientMatcher:
    def __init__(self, excel_path, column_mapping):
        self.excel_path = excel_path
        self.column_mapping = column_mapping
        self.recipients = []
        self._load_recipients()

    def _load_recipients(self):
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active

        key_col = self.column_mapping.get_key_column()
        if not key_col:
            raise ValueError("请先设置关键字段列")

        for row in ws.iter_rows(min_row=2, values_only=True):
            key = self.column_mapping.get_data_for_field(row, 'key')
            name = self.column_mapping.get_data_for_field(row, 'name')
            email = self.column_mapping.get_data_for_field(row, 'email')

            if key:
                self.recipients.append({
                    'key': key,
                    'name': name,
                    'email': email,
                    'raw_data': row
                })
        wb.close()

    def find_by_key(self, key):
        for r in self.recipients:
            if r['key'] == key:
                return r
        return None

    def match_files(self, folder_path):
        results = []
        if not os.path.exists(folder_path):
            return results

        for filename in os.listdir(folder_path):
            if os.path.isfile(os.path.join(folder_path, filename)):
                key = extract_key_from_filename(filename)
                if not key:
                    results.append({
                        'file': filename,
                        'filepath': os.path.join(folder_path, filename),
                        'status': 'invalid_filename',
                        'error': '文件名不符合格式（应为数字开头）'
                    })
                    continue

                recipient = self.find_by_key(key)
                if recipient:
                    results.append({
                        'file': filename,
                        'filepath': os.path.join(folder_path, filename),
                        'status': 'matched',
                        'recipient': recipient
                    })
                else:
                    results.append({
                        'file': filename,
                        'filepath': os.path.join(folder_path, filename),
                        'status': 'not_found',
                        'key': key,
                        'error': f'学号 {key} 在Excel中未找到'
                    })
        return results


class EmailSenderGUI:
    def __init__(self):
        if USE_CUSTOMTK:
            ctk.set_appearance_mode("light")
            ctk.set_default_color_theme("blue")

        self.root = tk.Tk() if not USE_CUSTOMTK else ctk.CTk()
        self.root.title("批量发送邮件工具")
        self.root.geometry("900x700")

        self.current_step = 1
        self.folder_path = ""
        self.excel_path = ""
        self.column_mapping = ColumnMapping()
        self.recipient_matcher = None
        self.match_results = []
        self.sending = False
        self.paused = False
        self.stop_event = threading.Event()

        self.smtp_config = {
            'server': '',
            'port': 587,
            'email': '',
            'password': '',
            'use_tls': True
        }

        self.send_interval = 3
        self.max_retries = 3

        # Tab键导航支持
        def move_focus_on_tab(event):
            event.widget.tk_focusNext().focus_set()
            return 'break'

        self._move_focus_on_tab = move_focus_on_tab

        self.setup_responsive_ui()

    def get_screen_size(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        return screen_width, screen_height

    def setup_responsive_ui(self):
        screen_width, screen_height = self.get_screen_size()
        
        min_width = min(900, screen_width - 50)
        min_height = min(700, screen_height - 100)
        
        self.root.minsize(800, 600)
        
        if screen_width < 900 or screen_height < 700:
            base_geometry = f"{screen_width - 50}x{screen_height - 100}"
            self.root.geometry(base_geometry)
        else:
            self.root.geometry("900x700")
        
        container = tk.Frame(self.root) if not USE_CUSTOMTK else ctk.CTkFrame(self.root)
        container.pack(fill='both', expand=True)
        
        if screen_width < 900 or screen_height < 700:
            self.canvas = tk.Canvas(container) if not USE_CUSTOMTK else None
            scrollbar_y = tk.Scrollbar(container, orient='vertical', command=self.canvas.yview) if not USE_CUSTOMTK else None
            scrollbar_x = tk.Scrollbar(container, orient='horizontal', command=self.canvas.xview) if not USE_CUSTOMTK else None
            
            if not USE_CUSTOMTK and self.canvas:
                self.canvas.configure(xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)
                
                def on_frame_configure(event):
                    self.canvas.configure(scrollregion=self.canvas.bbox('all'))
                
                self.main_frame = tk.Frame(self.canvas) if not USE_CUSTOMTK else None
                self.canvas.create_window((0, 0), window=self.main_frame, anchor='nw')
                self.main_frame.bind("<Configure>", on_frame_configure)
                
                self.canvas.pack(side='left', fill='both', expand=True)
                scrollbar_y.pack(side='right', fill='y')
                scrollbar_x.pack(side='bottom', fill='x')
            else:
                self.main_frame = container
        else:
            self.main_frame = container

        self.setup_ui()

    def setup_ui(self):
        main_frame = tk.Frame(self.main_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.main_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        step_frame = tk.Frame(main_frame) if not USE_CUSTOMTK else ctk.CTkFrame(main_frame, height=60)
        step_frame.pack(fill='x', pady=(0, 10))
        step_frame.pack_propagate(False) if USE_CUSTOMTK else None

        self.step_buttons = []
        steps = ['1.数据配置', '2.邮件内容', '3.发送设置', '4.发送预览']
        for i, step_text in enumerate(steps, 1):
            btn = tk.Button(step_frame, text=step_text, command=lambda idx=i: self.go_to_step(idx),
                          width=12, height=2) if not USE_CUSTOMTK else ctk.CTkButton(
                step_frame, text=step_text, command=lambda idx=i: self.go_to_step(idx), width=12, height=40)
            btn.pack(side='left', padx=5)
            self.step_buttons.append(btn)

        self.content_frame = tk.Frame(main_frame) if not USE_CUSTOMTK else ctk.CTkFrame(main_frame)
        self.content_frame.pack(fill='both', expand=True)

        self.step1_frame = None
        self.step2_frame = None
        self.step3_frame = None
        self.step4_frame = None

        btn_frame = tk.Frame(main_frame) if not USE_CUSTOMTK else ctk.CTkFrame(main_frame, height=50)
        btn_frame.pack(fill='x', pady=(10, 0))
        btn_frame.pack_propagate(False) if USE_CUSTOMTK else None

        self.btn_prev = tk.Button(btn_frame, text="上一步", command=self.prev_step, width=10) if not USE_CUSTOMTK else ctk.CTkButton(
            btn_frame, text="上一步", command=self.prev_step, width=10)
        self.btn_prev.pack(side='left', padx=5)
        self.btn_prev.pack_forget()

        self.btn_next = tk.Button(btn_frame, text="下一步", command=self.next_step, width=10) if not USE_CUSTOMTK else ctk.CTkButton(
            btn_frame, text="下一步", command=self.next_step, width=10)
        self.btn_next.pack(side='left', padx=5)

        self.btn_send = tk.Button(btn_frame, text="开始发送", command=self.start_send, width=12,
                               bg='green', fg='white') if not USE_CUSTOMTK else ctk.CTkButton(
            btn_frame, text="开始发送", command=self.start_send, width=12, fg_color='green', hover_color='darkgreen')
        self.btn_send.pack(side='left', padx=5)
        self.btn_send.pack_forget()

        self.create_step1()
        self.create_step2()
        self.create_step3()
        self.create_step4()

        self.show_step(1)

        # 键盘导航支持
        self.root.bind('<Control-Return>', lambda e: self.start_send() if self.current_step == 4 else None)
        self.root.bind('<Control-q>', lambda e: self.on_closing())
        self.root.bind('<Escape>', lambda e: None)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_step1(self):
        self.step1_frame = tk.Frame(self.content_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.content_frame)

        tk.Label(self.step1_frame, text="步骤1: 数据配置", font=('Arial', 14, 'bold')).pack(anchor='w', pady=(0, 15))

        tk.Label(self.step1_frame, text="附件文件夹:").pack(anchor='w')
        folder_frame = tk.Frame(self.step1_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.step1_frame)
        folder_frame.pack(fill='x', pady=(5, 10))
        self.entry_folder = tk.Entry(folder_frame, width=60) if not USE_CUSTOMTK else ctk.CTkEntry(folder_frame, width=60)
        self.entry_folder.pack(side='left', fill='x', expand=True)
        tk.Button(folder_frame, text="浏览...", command=self.browse_folder, width=8).pack(side='left', padx=(5, 0)) if not USE_CUSTOMTK else ctk.CTkButton(
            folder_frame, text="浏览...", command=self.browse_folder, width=8).pack(side='left', padx=(5, 0))

        tk.Label(self.step1_frame, text="Excel文件:").pack(anchor='w')
        excel_frame = tk.Frame(self.step1_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.step1_frame)
        excel_frame.pack(fill='x', pady=(5, 10))
        self.entry_excel = tk.Entry(excel_frame, width=60) if not USE_CUSTOMTK else ctk.CTkEntry(excel_frame, width=60)
        self.entry_excel.pack(side='left', fill='x', expand=True)
        tk.Button(excel_frame, text="浏览...", command=self.browse_excel, width=8).pack(side='left', padx=(5, 0)) if not USE_CUSTOMTK else ctk.CTkButton(
            excel_frame, text="浏览...", command=self.browse_excel, width=8).pack(side='left', padx=(5, 0))

        sep = tk.Frame(self.step1_frame, height=2, bd=1, relief='sunken') if not USE_CUSTOMTK else ctk.CTkFrame(self.step1_frame, height=2)
        sep.pack(fill='x', pady=15)

        tk.Label(self.step1_frame, text="列映射配置（请为每列选择用途）:").pack(anchor='w')
        tk.Label(self.step1_frame, text="提示：关键字段用于匹配文件名中的编号", fg='gray').pack(anchor='w')

        self.mapping_frame = tk.Frame(self.step1_frame) if not USE_CUSTOMTK else ctk.CTkScrollableFrame(self.step1_frame, height=200)
        self.mapping_frame.pack(fill='both', expand=True, pady=10)

        self.mapping_widgets = []

        sep2 = tk.Frame(self.step1_frame, height=2, bd=1, relief='sunken') if not USE_CUSTOMTK else ctk.CTkFrame(self.step1_frame, height=2)
        sep2.pack(fill='x', pady=15)

        tk.Label(self.step1_frame, text="匹配预览:").pack(anchor='w')

        self.preview_frame = tk.Frame(self.step1_frame) if not USE_CUSTOMTK else ctk.CTkScrollableFrame(self.step1_frame, height=150)
        self.preview_frame.pack(fill='both', expand=True, pady=5)

        self.btn_load_preview = tk.Button(self.step1_frame, text="加载预览", command=self.load_preview,
                                       bg=get_color('lightblue')) if not USE_CUSTOMTK else ctk.CTkButton(
            self.step1_frame, text="加载预览", command=self.load_preview, fg_color=get_color('lightblue'), text_color='black')
        self.btn_load_preview.pack(pady=5)
        
        self.preview_hint_label = tk.Label(self.step1_frame, text="[!] 重要：必须点击'加载预览'确保有匹配收件人后才能继续！",
                                          fg=get_color('red'), font=('Arial', 9, 'bold'))
        if not USE_CUSTOMTK:
            self.preview_hint_label.pack(pady=(0, 5))
        else:
            self.preview_hint_label.configure(text_color=get_color('red'))

    def create_step2(self):
        self.step2_frame = tk.Frame(self.content_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.content_frame)

        tk.Label(self.step2_frame, text="步骤2: 邮件内容", font=('Arial', 14, 'bold')).pack(anchor='w', pady=(0, 15))

        tk.Label(self.step2_frame, text="邮件主题:").pack(anchor='w')
        self.entry_subject = tk.Entry(self.step2_frame, width=70) if not USE_CUSTOMTK else ctk.CTkEntry(self.step2_frame, width=70)
        self.entry_subject.pack(fill='x', pady=(5, 5))
        
        tk.Label(self.step2_frame, text="主题变量:", font=('Arial', 9)).pack(anchor='w')
        subject_var_frame = tk.Frame(self.step2_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.step2_frame)
        subject_var_frame.pack(pady=(2, 10))
        self.btn_insert_subject_name = tk.Button(subject_var_frame, text="【姓名】", command=lambda: self.insert_var_to_subject('name'),
                                               width=8).pack(side='left', padx=2)
        self.btn_insert_subject_email = tk.Button(subject_var_frame, text="【邮箱】", command=lambda: self.insert_var_to_subject('email'),
                                               width=8).pack(side='left', padx=2)

        tk.Label(self.step2_frame, text="插入变量(正文):").pack(anchor='w')
        var_frame = tk.Frame(self.step2_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.step2_frame)
        var_frame.pack(pady=5)
        self.btn_insert_key = tk.Button(var_frame, text="【姓名】", command=lambda: self.insert_var('name'),
                                      width=8).pack(side='left', padx=2)
        self.btn_insert_email = tk.Button(var_frame, text="【邮箱】", command=lambda: self.insert_var('email'),
                                       width=8).pack(side='left', padx=2)

        format_label_frame = tk.LabelFrame(self.step2_frame, text="文本格式化工具", padx=5, pady=5)
        format_label_frame.pack(fill='x', pady=(10, 5))

        # 第一行：对齐方式
        align_toolbar = tk.Frame(format_label_frame) if not USE_CUSTOMTK else ctk.CTkFrame(format_label_frame)
        align_toolbar.pack(fill='x', pady=2)

        tk.Label(align_toolbar, text="对齐方式:").pack(side='left', padx=(0, 5))

        self.btn_align_left = tk.Button(align_toolbar, text="左对齐", command=self.format_align_left,
                                        width=8, relief='raised')
        if not USE_CUSTOMTK:
            self.btn_align_left.pack(side='left', padx=2)
        else:
            self.btn_align_left = ctk.CTkButton(align_toolbar, text="左对齐", command=self.format_align_left, width=8)
            self.btn_align_left.pack(side='left', padx=2)

        self.btn_align_right = tk.Button(align_toolbar, text="右对齐", command=self.format_align_right,
                                         width=8, relief='raised')
        if not USE_CUSTOMTK:
            self.btn_align_right.pack(side='left', padx=2)
        else:
            self.btn_align_right = ctk.CTkButton(align_toolbar, text="右对齐", command=self.format_align_right, width=8)
            self.btn_align_right.pack(side='left', padx=2)

        # 第二行：样式
        style_toolbar = tk.Frame(format_label_frame) if not USE_CUSTOMTK else ctk.CTkFrame(format_label_frame)
        style_toolbar.pack(fill='x', pady=2)

        tk.Label(style_toolbar, text="样式:").pack(side='left', padx=(0, 5))

        self.btn_bold = tk.Button(style_toolbar, text="粗体", command=self.format_bold,
                                  width=8, relief='raised')
        if not USE_CUSTOMTK:
            self.btn_bold.pack(side='left', padx=2)
        else:
            self.btn_bold = ctk.CTkButton(style_toolbar, text="粗体", command=self.format_bold, width=8)
            self.btn_bold.pack(side='left', padx=2)

        # 第三行：缩进和清除
        indent_toolbar = tk.Frame(format_label_frame) if not USE_CUSTOMTK else ctk.CTkFrame(format_label_frame)
        indent_toolbar.pack(fill='x', pady=2)

        tk.Label(indent_toolbar, text="缩进:").pack(side='left', padx=(0, 5))

        self.btn_indent = tk.Button(indent_toolbar, text="首行缩进", command=self.format_first_line_indent,
                                    width=10, relief='raised')
        if not USE_CUSTOMTK:
            self.btn_indent.pack(side='left', padx=2)
        else:
            self.btn_indent = ctk.CTkButton(indent_toolbar, text="首行缩进", command=self.format_first_line_indent, width=10)
            self.btn_indent.pack(side='left', padx=2)

        self.btn_clear_format = tk.Button(indent_toolbar, text="清除格式", command=self.clear_formatting,
                                          width=10, relief='raised')
        if not USE_CUSTOMTK:
            self.btn_clear_format.pack(side='left', padx=2)
        else:
            self.btn_clear_format = ctk.CTkButton(indent_toolbar, text="清除格式", command=self.clear_formatting, width=10)
            self.btn_clear_format.pack(side='left', padx=2)

        tk.Label(self.step2_frame, text="邮件正文:").pack(anchor='w', pady=(10, 0))
        
        text_frame = tk.Frame(self.step2_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.step2_frame)
        text_frame.pack(fill='both', expand=True, pady=5)
        
        self.text_body = scrolledtext.ScrolledText(text_frame, width=80, height=15,
                                                font=('Courier', 10)) if not USE_CUSTOMTK else ctk.CTkTextbox(
            text_frame, width=80, height=15, font=('Courier', 10))
        self.text_body.pack(fill='both', expand=True)

        # Tab键导航支持：文本区域按Tab移动焦点
        self.text_body.bind('<Tab>', self._move_focus_on_tab)
        if USE_CUSTOMTK:
            self.text_body.bind('<Tab>', self._move_focus_on_tab)

        self.text_body.tag_configure('bold', font=('Courier', 10, 'bold'))
        self.text_body.tag_configure('left_align', justify='left')
        self.text_body.tag_configure('right_align', justify='right')
        self.text_body.tag_configure('indented', lmargin1=40, lmargin2=40)

        self.default_body = """您好，【姓名】同学：

您好！附件中是您的学习资料，请查收。

如有疑问，请随时联系我们。

祝学习顺利！"""

        if not USE_CUSTOMTK:
            self.text_body.insert('1.0', self.default_body)
        else:
            self.text_body.insert('1.0', self.default_body)

    def create_step3(self):
        self.step3_frame = tk.Frame(self.content_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.content_frame)

        tk.Label(self.step3_frame, text="步骤3: 发送设置", font=('Arial', 14, 'bold')).pack(anchor='w', pady=(0, 15))

        config_frame = tk.LabelFrame(self.step3_frame, text="SMTP配置", padx=10, pady=10)
        config_frame.pack(fill='x', pady=(0, 15))

        row1 = tk.Frame(config_frame)
        row1.pack(fill='x', pady=5)
        tk.Label(row1, text="服务器:", width=10, anchor='w').pack(side='left')
        self.entry_server = tk.Entry(row1, width=40) if not USE_CUSTOMTK else ctk.CTkEntry(row1, width=40)
        self.entry_server.pack(side='left', fill='x', expand=True)
        self.entry_server.insert(0, 'smtp.example.com')

        row2 = tk.Frame(config_frame)
        row2.pack(fill='x', pady=5)
        tk.Label(row2, text="端口:", width=10, anchor='w').pack(side='left')
        self.entry_port = tk.Entry(row2, width=15) if not USE_CUSTOMTK else ctk.CTkEntry(row2, width=15)
        self.entry_port.pack(side='left')
        self.entry_port.insert(0, '587')
        self.var_tls = tk.BooleanVar(value=True)
        tk.Checkbutton(row2, text="使用TLS", variable=self.var_tls).pack(side='left', padx=20)

        row3 = tk.Frame(config_frame)
        row3.pack(fill='x', pady=5)
        tk.Label(row3, text="邮箱:", width=10, anchor='w').pack(side='left')
        self.entry_sender = tk.Entry(row3, width=40) if not USE_CUSTOMTK else ctk.CTkEntry(row3, width=40)
        self.entry_sender.pack(side='left', fill='x', expand=True)

        row4 = tk.Frame(config_frame)
        row4.pack(fill='x', pady=5)
        tk.Label(row4, text="授权码:", width=10, anchor='w').pack(side='left')
        self.entry_auth = tk.Entry(row4, width=40, show='*') if not USE_CUSTOMTK else ctk.CTkEntry(row4, width=40, show='*')
        self.entry_auth.pack(side='left', fill='x', expand=True)

        tk.Button(config_frame, text="测试连接", command=self.test_connection,
                 bg='lightgreen').pack(pady=10)

        options_frame = tk.LabelFrame(self.step3_frame, text="发送选项", padx=10, pady=10)
        options_frame.pack(fill='x')

        row5 = tk.Frame(options_frame)
        row5.pack(fill='x', pady=5)
        tk.Label(row5, text="发送间隔:", width=10, anchor='w').pack(side='left')
        if not USE_CUSTOMTK:
            self.var_interval = tk.StringVar(value='3')
            vcmd_interval = (self.root.register(self.validate_interval), '%P')
            self.entry_interval = tk.Entry(row5, width=10, textvariable=self.var_interval, validate='key', validatecommand=vcmd_interval)
            self.entry_interval.pack(side='left')
            self.var_interval.trace_add('write', self.on_interval_change)
        else:
            self.entry_interval = ctk.CTkEntry(row5, width=10)
            self.entry_interval.pack(side='left')
            self.entry_interval.insert(0, '3')
            self.entry_interval.bind('<KeyRelease>', self.on_ctk_interval_change)
        tk.Label(row5, text="秒（建议3秒以上）").pack(side='left', padx=10)

        row6 = tk.Frame(options_frame)
        row6.pack(fill='x', pady=5)
        tk.Label(row6, text="最大重试:", width=10, anchor='w').pack(side='left')
        if not USE_CUSTOMTK:
            self.var_retries = tk.StringVar(value='3')
            vcmd_retries = (self.root.register(self.validate_retries), '%P')
            self.entry_retries = tk.Entry(row6, width=10, textvariable=self.var_retries, validate='key', validatecommand=vcmd_retries)
            self.entry_retries.pack(side='left')
            self.var_retries.trace_add('write', self.on_retries_change)
        else:
            self.entry_retries = ctk.CTkEntry(row6, width=10)
            self.entry_retries.pack(side='left')
            self.entry_retries.insert(0, '3')
            self.entry_retries.bind('<KeyRelease>', self.on_ctk_retries_change)
        tk.Label(row6, text="次").pack(side='left', padx=10)

    def create_step4(self):
        self.step4_frame = tk.Frame(self.content_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.content_frame)

        tk.Label(self.step4_frame, text="步骤4: 发送预览与执行", font=('Arial', 14, 'bold')).pack(anchor='w', pady=(0, 15))

        self.label_stats = tk.Label(self.step4_frame, text="总计: 0  成功: 0  失败: 0",
                                  font=('Arial', 12))
        self.label_stats.pack(anchor='w')

        self.progress_bar = ttk.Progressbar(self.step4_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=10)

        log_frame = tk.LabelFrame(self.step4_frame, text="发送日志", padx=5, pady=5)
        log_frame.pack(fill='both', expand=True)
        
        log_text_frame = tk.Frame(log_frame) if not USE_CUSTOMTK else ctk.CTkFrame(log_frame)
        log_text_frame.pack(fill='both', expand=True)
        
        self.text_log = scrolledtext.ScrolledText(log_text_frame, width=80, height=20,
                                                     font=('Courier', 9))
        self.text_log.pack(fill='both', expand=True)
        self.text_log.bind('<Tab>', self._move_focus_on_tab)

        btn_log_frame = tk.Frame(self.step4_frame)
        btn_log_frame.pack(pady=10)

        self.btn_pause = tk.Button(btn_log_frame, text="暂停", command=self.toggle_pause,
                                  width=10, state='disabled') if not USE_CUSTOMTK else ctk.CTkButton(
            btn_log_frame, text="暂停", command=self.toggle_pause, width=10, state='disabled')
        self.btn_pause.pack(side='left', padx=5)

        self.btn_stop = tk.Button(btn_log_frame, text="停止", command=self.stop_send,
                                 width=10, state='disabled', bg='red', fg='white') if not USE_CUSTOMTK else ctk.CTkButton(
            btn_log_frame, text="停止", command=self.stop_send, width=10, state='disabled', fg_color='red')
        self.btn_stop.pack(side='left', padx=5)

        self.btn_export = tk.Button(btn_log_frame, text="导出失败记录", command=self.export_failed,
                                   width=12) if not USE_CUSTOMTK else ctk.CTkButton(
            btn_log_frame, text="导出失败记录", command=self.export_failed, width=12)
        self.btn_export.pack(side='left', padx=5)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="选择附件文件夹")
        if folder:
            self.folder_path = folder
            self.entry_folder.delete(0, 'end')
            self.entry_folder.insert(0, folder)

    def browse_excel(self):
        file = filedialog.askopenfilename(title="选择Excel文件",
                                          filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if file:
            self.excel_path = file
            self.entry_excel.delete(0, 'end')
            self.entry_excel.insert(0, file)
            self.load_columns()

    def load_columns(self):
        for widget in self.mapping_widgets:
            widget.destroy()
        self.mapping_widgets = []

        if not self.excel_path:
            return

        try:
            self.column_mapping.load_from_excel(self.excel_path)

            for col_info in self.column_mapping.columns:
                row = tk.Frame(self.mapping_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.mapping_frame)
                row.pack(fill='x', pady=2)
                self.mapping_widgets.append(row)

                letter = col_info['letter']
                name = col_info['name']
                suggested = col_info['suggested']

                tk.Label(row, text=f"{letter} - {name}", width=20, anchor='w').pack(side='left')

                var = tk.StringVar(value=suggested)
                combo = ttk.Combobox(row, values=FIELD_TYPE_KEYS, textvariable=var,
                                    width=15, state='readonly')
                combo.pack(side='left', padx=10)
                self.mapping_widgets.append(combo)

                def on_change(event, letter=letter, var=var):
                    self.column_mapping.set_mapping(letter, var.get())
                combo.bind('<<ComboboxSelected>>', on_change)

                self.column_mapping.mapping[letter] = suggested

        except Exception as e:
            messagebox.showerror("错误", f"读取Excel失败: {str(e)}")

    def load_preview(self):
        self.btn_load_preview.config(state='disabled', text='加载中...')
        self.btn_load_preview.update_idletasks()
        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        if not self.excel_path or not self.folder_path:
            self.btn_load_preview.config(state='normal', text='加载预览')
            messagebox.showwarning("警告", "请先选择Excel文件和附件文件夹")
            return

        try:
            self.column_mapping.load_from_excel(self.excel_path)
            self.recipient_matcher = RecipientMatcher(self.excel_path, self.column_mapping)
            self.match_results = self.recipient_matcher.match_files(self.folder_path)

            headers = ['附件文件', '收件人', '状态']
            header_row = tk.Frame(self.preview_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.preview_frame)
            header_row.pack(fill='x')
            for h in headers:
                tk.Label(header_row, text=h, anchor='w', font=('Arial', 9, 'bold')).pack(side='left', padx=5, fill='x', expand=True)

            matched = 0
            not_found = 0
            invalid = 0

            for result in self.match_results:
                row = tk.Frame(self.preview_frame) if not USE_CUSTOMTK else ctk.CTkFrame(self.preview_frame)
                row.pack(fill='x')

                status = result['status']
                if status == 'matched':
                    matched += 1
                    recipient = result['recipient']
                    status_text = f"{recipient['name']} <{recipient['email']}> [OK]"
                    color = 'green'
                elif status == 'not_found':
                    not_found += 1
                    status_text = result['error']
                    color = 'orange'
                else:
                    invalid += 1
                    status_text = result['error']
                    color = 'red'

                tk.Label(row, text=result['file'], anchor='w').pack(side='left', padx=5, fill='x', expand=True)

                if status == 'matched':
                    recipient = result['recipient']
                    tk.Label(row, text=f"{recipient['name']} <{recipient['email']}>",
                             anchor='w').pack(side='left', padx=5, fill='x', expand=True)
                else:
                    tk.Label(row, text="", anchor='w').pack(side='left', padx=5, fill='x', expand=True)

                label_color = get_color(color)
                if USE_CUSTOMTK:
                    tk.Label(row, text=status_text, anchor='w').pack(side='left', padx=5, fill='x', expand=True)
                    row.winfo_children()[-1].configure(text_color=label_color)
                else:
                    tk.Label(row, text=status_text, fg=label_color, anchor='w').pack(side='left', padx=5, fill='x', expand=True)

            summary = f"总计: {len(self.match_results)}  匹配成功: {matched}  未找到: {not_found}  无效: {invalid}"
            summary_label = tk.Label(self.preview_frame, text=summary, font=('Arial', 9, 'bold'),
            fg=get_color('blue') if not USE_CUSTOMTK else 'black')
            summary_label.pack(pady=5)
            
            if matched > 0:
                self.preview_hint_label.config(text=f"[OK] 预览已加载！找到 {matched} 个匹配收件人，可以继续操作",
                                             fg=get_color('green'))
                if not USE_CUSTOMTK:
                    self.preview_hint_label.config(fg=get_color('green'))
                else:
                    self.preview_hint_label.configure(text_color=get_color('green'))
            else:
                self.preview_hint_label.config(text="[!] 没有找到匹配的收件人！请检查文件名和Excel数据",
                                             fg=get_color('orange'))
                if not USE_CUSTOMTK:
                    self.preview_hint_label.config(fg=get_color('orange'))
                else:
                    self.preview_hint_label.configure(text_color=get_color('orange'))

            self.btn_load_preview.config(state='normal', text='加载预览')
        except Exception as e:
            self.btn_load_preview.config(state='normal', text='加载预览')
            messagebox.showerror("错误", f"预览失败: {str(e)}")

    def insert_var(self, var_type):
        var_map = {'name': '【姓名】', 'email': '【邮箱】'}
        var_text = var_map.get(var_type, '【】')

        try:
            if not USE_CUSTOMTK:
                self.text_body.insert('insert', var_text)
            else:
                self.text_body.insert('insert', var_text)
        except Exception as e:
            logging.error(f"insert_var error: {e}")

    def insert_var_to_subject(self, var_type):
        var_map = {'name': '【姓名】', 'email': '【邮箱】'}
        var_text = var_map.get(var_type, '【】')

        try:
            current_pos = self.entry_subject.index('insert')
            if current_pos == 'end':
                current_pos = len(self.entry_subject.get())
            self.entry_subject.insert(current_pos, var_text)
        except Exception as e:
            logging.error(f"insert_var_to_subject error: {e}")

    def format_align_left(self):
        try:
            if self.text_body.tag_ranges('sel'):
                start, end = self.text_body.tag_ranges('sel')
                self.text_body.tag_remove('right_align', start, end)
                self.text_body.tag_add('left_align', start, end)
            else:
                current_line = self.text_body.index('insert')
                line_start = current_line.split('.')[0] + '.0'
                line_end = current_line.split('.')[0] + '.end'
                self.text_body.tag_remove('right_align', line_start, line_end)
                self.text_body.tag_add('left_align', line_start, line_end)
        except Exception as e:
            messagebox.showwarning("提示", "请选择要格式化的文本或光标所在行")

    def format_align_right(self):
        try:
            if self.text_body.tag_ranges('sel'):
                start, end = self.text_body.tag_ranges('sel')
                self.text_body.tag_remove('left_align', start, end)
                self.text_body.tag_add('right_align', start, end)
            else:
                current_line = self.text_body.index('insert')
                line_start = current_line.split('.')[0] + '.0'
                line_end = current_line.split('.')[0] + '.end'
                self.text_body.tag_remove('left_align', line_start, line_end)
                self.text_body.tag_add('right_align', line_start, line_end)
        except Exception as e:
            messagebox.showwarning("提示", "请选择要格式化的文本或光标所在行")

    def format_bold(self):
        try:
            if self.text_body.tag_ranges('sel'):
                start, end = self.text_body.tag_ranges('sel')
                current_tags = self.text_body.tag_names(start)
                if 'bold' in current_tags:
                    self.text_body.tag_remove('bold', start, end)
                else:
                    self.text_body.tag_add('bold', start, end)
            else:
                messagebox.showwarning("提示", "请先选择要加粗的文本")
        except Exception as e:
            messagebox.showwarning("提示", "请先选择要加粗的文本")

    def format_first_line_indent(self):
        try:
            if self.text_body.tag_ranges('sel'):
                start, end = self.text_body.tag_ranges('sel')
                current_tags = self.text_body.tag_names(start)
                if 'indented' in current_tags:
                    self.text_body.tag_remove('indented', start, end)
                else:
                    self.text_body.tag_add('indented', start, end)
            else:
                current_line = self.text_body.index('insert')
                line_start = current_line.split('.')[0] + '.0'
                line_end = current_line.split('.')[0] + '.end'
                current_tags = self.text_body.tag_names(line_start)
                if 'indented' in current_tags:
                    self.text_body.tag_remove('indented', line_start, line_end)
                else:
                    self.text_body.tag_add('indented', line_start, line_end)
        except Exception as e:
            messagebox.showwarning("提示", "请选择要缩进的文本或光标所在行")

    def clear_formatting(self):
        try:
            if self.text_body.tag_ranges('sel'):
                start, end = self.text_body.tag_ranges('sel')
                for tag in ['bold', 'left_align', 'right_align', 'indented']:
                    self.text_body.tag_remove(tag, start, end)
            else:
                for tag in ['bold', 'left_align', 'right_align', 'indented']:
                    self.text_body.tag_remove(tag, '1.0', 'end')
                messagebox.showinfo("提示", "已清除所有格式")
        except Exception as e:
            pass

    def convert_to_html(self, text):
        lines = text.split('\n')
        html_lines = []
        
        for line in lines:
            if not line.strip():
                html_lines.append('<br>')
                continue
            
            line_html = line
            
            bold_ranges = []
            pos = '1.0'
            while True:
                next_range = self.text_body.tag_nextrange('bold', pos)
                if not next_range:
                    break
                bold_ranges.append(next_range)
                pos = next_range[1]
            
            align_right_ranges = []
            pos = '1.0'
            while True:
                next_range = self.text_body.tag_nextrange('right_align', pos)
                if not next_range:
                    break
                align_right_ranges.append(next_range)
                pos = next_range[1]
            
            indent_ranges = []
            pos = '1.0'
            while True:
                next_range = self.text_body.tag_nextrange('indented', pos)
                if not next_range:
                    break
                indent_ranges.append(next_range)
                pos = next_range[1]
            
            is_right_align = False
            is_indented = False
            
            html_lines.append(f'<p style="text-align: right;">{line_html}</p>' if is_right_align else 
                            f'<p style="text-indent: 2em;">{line_html}</p>' if is_indented else 
                            f'<p>{line_html}</p>')
        
        return '\n'.join(html_lines)

    def get_formatted_body_html(self):
        full_text = self.text_body.get('1.0', 'end')
        lines = full_text.split('\n')
        html_parts = []
        
        for i, line in enumerate(lines):
            if not line.strip():
                html_parts.append('<br>')
                continue
            
            line_start = f'{i+1}.0'
            line_end = f'{i+1}.end'
            
            chars = list(line)
            char_html = [''] * len(chars)
            
            bold_positions = set()
            right_align_positions = set()
            indent_positions = set()
            
            for j in range(len(chars)):
                pos = f'{i+1}.{j}'
                tags = self.text_body.tag_names(pos)
                if 'bold' in tags:
                    bold_positions.add(j)
                if 'right_align' in tags:
                    right_align_positions.add(j)
                if 'indented' in tags:
                    indent_positions.add(j)
            
            has_right_align = len(right_align_positions) > len(chars) // 2 if chars else False
            has_indent = len(indent_positions) > len(chars) // 2 if chars else False
            
            result_line = ''
            in_bold = False
            for j, char in enumerate(chars):
                is_bold = j in bold_positions
                if is_bold and not in_bold:
                    result_line += '<strong>'
                    in_bold = True
                elif not is_bold and in_bold:
                    result_line += '</strong>'
                    in_bold = False
                result_line += char if char != '<' else '&lt;' if char != '>' else '&gt;'
            
            if in_bold:
                result_line += '</strong>'
            
            style = ''
            if has_right_align:
                style = 'text-align: right;'
            elif has_indent:
                style = 'text-indent: 2em;'
            
            if style:
                html_parts.append(f'<p style="{style}">{result_line}</p>')
            else:
                html_parts.append(f'<p>{result_line}</p>')
        
        return '\n'.join(html_parts)

    def validate_interval(self, value):
        if value == '':
            return True
        if value.isdigit():
            return True
        return False

    def validate_retries(self, value):
        if value == '':
            return True
        if value.isdigit():
            return True
        return False

    def on_interval_change(self, *args):
        value = self.var_interval.get()
        if value and not value.isdigit():
            self.var_interval.set(''.join(c for c in value if c.isdigit()))

    def on_retries_change(self, *args):
        value = self.var_retries.get()
        if value and not value.isdigit():
            self.var_retries.set(''.join(c for c in value if c.isdigit()))

    def on_ctk_interval_change(self, event):
        value = self.entry_interval.get()
        if value and not value.isdigit():
            self.entry_interval.delete(0, 'end')
            self.entry_interval.insert(0, ''.join(c for c in value if c.isdigit()))

    def on_ctk_retries_change(self, event):
        value = self.entry_retries.get()
        if value and not value.isdigit():
            self.entry_retries.delete(0, 'end')
            self.entry_retries.insert(0, ''.join(c for c in value if c.isdigit()))

    def test_connection(self):
        server = self.entry_server.get().strip()
        port_str = self.entry_port.get().strip()
        email = self.entry_sender.get().strip()
        auth = self.entry_auth.get().strip()
        use_tls = self.var_tls.get()

        if not all([server, port_str, email, auth]):
            messagebox.showwarning("警告", "请填写完整的SMTP配置")
            return

        try:
            port = int(port_str)
            if not (1 <= port <= 65535):
                messagebox.showwarning("警告", "端口号必须在1-65535之间")
                return
        except ValueError:
            messagebox.showerror("错误", "端口必须是数字")
            return

        self.log_message("正在测试连接...")

        try:
            if use_tls:
                server_smtp = smtplib.SMTP(server, port, timeout=10)
                server_smtp.starttls()
            else:
                server_smtp = smtplib.SMTP(server, port, timeout=10)

            server_smtp.login(email, auth)
            server_smtp.quit()

            self.log_message("连接测试成功！")
            messagebox.showinfo("成功", "SMTP连接测试成功！")

            self.smtp_config = {
                'server': server,
                'port': port,
                'email': email,
                'password': auth,
                'use_tls': use_tls
            }

        except smtplib.SMTPAuthenticationError:
            self.log_message("认证失败，请检查邮箱和授权码")
            messagebox.showerror("错误", "认证失败，请检查邮箱和授权码")
        except Exception as e:
            self.log_message(f"连接失败: {str(e)}")
            messagebox.showerror("错误", f"连接失败: {str(e)}")

    def show_step(self, step):
        self.current_step = step

        for i, btn in enumerate(self.step_buttons, 1):
            if i == step:
                if not USE_CUSTOMTK:
                    btn.config(relief='sunken', bg='#4CAF50', fg='white')
                else:
                    btn.configure(fg_color='#2196F3', hover_color='#1976D2', text_color='white')
            else:
                if not USE_CUSTOMTK:
                    btn.config(relief='raised', bg='SystemButtonFace', fg='black')
                else:
                    btn.configure(fg_color='#E0E0E0', hover_color='#BDBDBD', text_color='black')

        for frame in [self.step1_frame, self.step2_frame, self.step3_frame, self.step4_frame]:
            if frame:
                frame.pack_forget()

        current = [self.step1_frame, self.step2_frame, self.step3_frame, self.step4_frame][step - 1]
        if current:
            current.pack(fill='both', expand=True)

        self.btn_prev.pack_forget() if step <= 1 else self.btn_prev.pack(side='left', padx=5)

        if step < 4:
            self.btn_next.pack(side='left', padx=5)
            self.btn_send.pack_forget()
        else:
            self.btn_next.pack_forget()
            self.btn_send.pack(side='left', padx=5)

    def go_to_step(self, step):
        if step <= self.current_step or self.validate_step(self.current_step):
            self.show_step(step)

    def validate_step(self, step):
        if step == 1:
            if not self.excel_path or not self.folder_path:
                messagebox.showwarning("警告", "请先选择Excel文件和附件文件夹")
                return False
            if not self.column_mapping.get_key_column():
                messagebox.showwarning("警告", "请设置关键字段列")
                return False
            matched_results = [r for r in self.match_results if r['status'] == 'matched']
            if not matched_results:
                messagebox.showwarning("警告", "请点击'加载预览'按钮，确保有收件人匹配成功！")
                return False
        elif step == 2:
            subject = self.entry_subject.get().strip()
            if not subject:
                messagebox.showwarning("警告", "请填写邮件主题")
                return False
        elif step == 3:
            if not self.smtp_config.get('password'):
                messagebox.showwarning("警告", "请点击'测试连接'按钮，验证SMTP配置是否正确！")
                return False
        return True

    def prev_step(self):
        if self.current_step > 1:
            self.show_step(self.current_step - 1)

    def next_step(self):
        if self.validate_step(self.current_step):
            self.show_step(self.current_step + 1)

    def start_send(self):
        if not self.validate_step(3):
            return
        
        matched_results = [r for r in self.match_results if r['status'] == 'matched']
        if not matched_results:
            messagebox.showwarning("警告", "没有匹配到任何收件人！\n请在步骤1中点击'加载预览'按钮，确保文件和收件人匹配成功。")
            return

        self.sending = True
        self.paused = False
        self.stop_event.clear()

        self.btn_pause.config(state='normal', text='暂停')
        self.btn_stop.config(state='normal')
        self.btn_send.config(state='disabled')

        self.text_log.delete('1.0', 'end')
        self.label_stats.config(text=f"总计: {len(matched_results)}  成功: 0  失败: 0")
        self.progress_bar['value'] = 0

        subject = self.entry_subject.get().strip()
        body = self.text_body.get('1.0', 'end').strip()

        self.send_interval = int(self.entry_interval.get() or 3)
        self.max_retries = int(self.entry_retries.get() or 3)

        thread = threading.Thread(target=self.send_emails, args=(subject, body))
        thread.daemon = True
        thread.start()

    def send_emails(self, subject, body):
        total = len(self.match_results)
        matched_results = [r for r in self.match_results if r['status'] == 'matched']

        success_count = 0
        fail_count = 0
        
        body_html_template = self.get_formatted_body_html()
        
        def check_stop():
            if self.stop_event.is_set():
                return True
            self.root.update() if hasattr(self, 'root') else None
            return False

        for i, result in enumerate(matched_results):
            if check_stop():
                self.log_message("发送已停止")
                break

            while self.paused and not self.stop_event.is_set():
                time.sleep(0.2)
            
            if check_stop():
                self.log_message("发送已停止")
                break

            recipient = result['recipient']
            filename = result['file']
            filepath = result['filepath']

            self.log_message(f"发送中: {recipient['name']} <{recipient['email']}>")

            personalized_subject = subject.replace('【姓名】', recipient['name'])
            personalized_body_plain = body.replace('【姓名】', recipient['name'])
            personalized_body_plain = personalized_body_plain.replace('【邮箱】', recipient['email'])
            
            personalized_body_html = body_html_template.replace('【姓名】', recipient['name'])
            personalized_body_html = personalized_body_html.replace('【邮箱】', recipient['email'])

            email_sent = False
            for retry in range(self.max_retries):
                if check_stop():
                    self.log_message("发送已停止")
                    break
                
                try:
                    self.send_single_email(
                        self.smtp_config['server'],
                        self.smtp_config['port'],
                        self.smtp_config['email'],
                        self.smtp_config['password'],
                        self.smtp_config['use_tls'],
                        recipient['email'],
                        personalized_subject,
                        personalized_body_plain,
                        personalized_body_html,
                        filepath
                    )
                    self.log_message(f"  [OK] 发送成功")
                    success_count += 1
                    email_sent = True
                    break
                except Exception as e:
                    if check_stop():
                        break
                    if retry < self.max_retries - 1:
                        self.log_message(f"  [RETRY] 重试中... ({retry + 1}/{self.max_retries})")
                        for wait_time in range(3):
                            if check_stop():
                                break
                            time.sleep(1)
                    else:
                        self.log_message(f"  [FAILED] 发送失败: {str(e)}")
                        fail_count += 1
            
            if check_stop():
                break

            progress = ((i + 1) / len(matched_results)) * 100
            self.root.after(0, lambda p=progress: self.progress_bar.config(value=p))
            self.root.after(0, lambda s=success_count, f=fail_count:
                          self.label_stats.config(text=f"总计: {len(matched_results)}  成功: {s}  失败: {f}"))

            if i < len(matched_results) - 1 and not self.stop_event.is_set():
                for wait_time in range(self.send_interval):
                    if check_stop():
                        break
                    time.sleep(1)

        self.root.after(0, self.on_send_complete)
        self.log_message(f"\n发送完成！成功: {success_count}, 失败: {fail_count}")

    def send_single_email(self, server, port, email, password, use_tls, recipient, subject, body_plain, body_html, attachment_path):
        msg = MIMEMultipart('alternative')
        msg['From'] = email
        msg['To'] = recipient
        msg['Subject'] = Header(subject, 'utf-8').encode()
        
        msg.attach(MIMEText(body_plain, 'plain', 'utf-8'))
        
        html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
p {{ margin: 0.5em 0; }}
</style>
</head>
<body>
{body_html}
</body>
</html>"""
        msg.attach(MIMEText(html_content, 'html', 'utf-8'))

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            msg.attach(part)

        if use_tls:
            server_conn = smtplib.SMTP(server, port, timeout=30)
            server_conn.starttls()
        else:
            server_conn = smtplib.SMTP(server, port, timeout=30)

        server_conn.login(email, password)
        server_conn.sendmail(email, [recipient], msg.as_string())
        server_conn.quit()

    def toggle_pause(self):
        self.paused = not self.paused
        if self.paused:
            self.btn_pause.config(text='继续')
            self.log_message("发送已暂停")
        else:
            self.btn_pause.config(text='暂停')
            self.log_message("发送继续")

    def stop_send(self):
        self.stop_event.set()
        self.log_message("正在停止...")
        self.root.after(100, self.check_stop_status)

    def check_stop_status(self):
        if self.stop_event.is_set():
            self.btn_stop.config(state='disabled')
            self.log_message("正在等待当前操作完成...")

    def on_send_complete(self):
        self.sending = False
        self.stop_event.clear()
        self.btn_pause.config(state='disabled', text='暂停')
        self.btn_stop.config(state='disabled')
        self.btn_send.config(state='normal')
        self.log_message("发送操作已结束")

    def log_message(self, msg):
        timestamp = time.strftime('%H:%M:%S')
        self.text_log.insert('end', f"[{timestamp}] {msg}\n")
        self.text_log.see('end')

    def export_failed(self):
        matched_results = [r for r in self.match_results if r['status'] == 'matched']
        failed_results = [r for r in self.match_results if r['status'] != 'matched']

        if not failed_results:
            messagebox.showinfo("提示", "没有失败记录需要导出")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存失败记录",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"failed_records_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not file_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "发送失败记录"

            headers = ['序号', '文件名', '学号', '姓名', '邮箱', '失败原因']
            ws.append(headers)

            for i, result in enumerate(failed_results, 1):
                key = result.get('key', '')
                recipient = result.get('recipient', {})
                ws.append([
                    i,
                    result['file'],
                    key,
                    recipient.get('name', ''),
                    recipient.get('email', ''),
                    result.get('error', '')
                ])

            for col in range(1, len(headers) + 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20

            wb.save(file_path)
            messagebox.showinfo("成功", f"失败记录已保存至:\n{file_path}")

        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")

    def on_closing(self):
        if self.sending:
            if messagebox.askyesno("确认", "正在发送中，确定要退出吗？"):
                self.stop_event.set()
                self.root.after(500, self.root.destroy)
        else:
            self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = EmailSenderGUI()
    app.run()
