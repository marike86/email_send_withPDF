import os
import sys
import json
import time
import re
import argparse
import getpass
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
import smtplib
import openpyxl

try:
    from progress import PowerShellProgress, ColoredOutput, is_powershell
except ImportError:
    import importlib.util
    spec = importlib.util.spec_from_file_location("progress", os.path.join(os.path.dirname(__file__), "progress.py"))
    progress_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(progress_module)
    PowerShellProgress = progress_module.PowerShellProgress
    ColoredOutput = progress_module.ColoredOutput
    is_powershell = progress_module.is_powershell


DEFAULT_CONFIG = {
    # 【SMTP服务器配置】
    "smtp_server": "smtp.gmail.com",      # SMTP服务器地址，如: smtp.gmail.com, smtp.qq.com, smtp.sdu.edu.cn
    "smtp_port": 587,                      # SMTP端口，常用: 587(TLS), 465(SSL), 25
    "use_tls": True,                         # 是否使用TLS加密

    # 【文件路径配置】
    "email_template": "email_template.txt", # 邮件模板文件路径
    "student_info_excel": "student_info.xlsx", # 学生信息Excel文件路径

    # 【发送配置】
    "send_interval": 3,      # 发送间隔秒数，建议3-5秒，避免被识别为垃圾邮件
    "max_retries": 3,        # 失败重试次数
    "retry_interval": 10     # 重试间隔秒数
}

STATUS_FILE = "status.json"


def load_config(config_path="config.json"):
    config = DEFAULT_CONFIG.copy()
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                config.update(loaded)
        except Exception as e:
            print(f"警告：读取配置文件失败: {e}")

    for key in ['smtp_server', 'smtp_port', 'use_tls', 'sender_email', 'sender_password', 'email_template', 'student_info_excel']:
        env_key = f"SMTP_{key.upper()}" if key.startswith('smtp') else key.upper()
        env_val = os.getenv(env_key)
        if env_val is not None:
            if key == 'smtp_port':
                config[key] = int(env_val)
            elif key == 'use_tls':
                config[key] = env_val.lower() == 'true'
            else:
                config[key] = env_val

    for key in ['email_template', 'student_info_excel']:
        if key in config and config[key]:
            config[key] = clean_path(config[key])

    return config


def clean_path(path):
    """清理路径中的引号和空格"""
    if not path:
        return path
    path = str(path).strip()
    while path.startswith('"') and path.endswith('"'):
        path = path[1:-1]
    while path.startswith("'"):
        path = path[1:]
    while path.endswith("'"):
        path = path[:-1]
    return path.strip()


def save_config(config, config_path="config.json"):
    save_config_data = config.copy()
    for key in ['email_template', 'student_info_excel']:
        if key in save_config_data and save_config_data[key]:
            save_config_data[key] = clean_path(save_config_data[key])
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(save_config_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"警告：保存配置文件失败: {e}")


def get_all_config_from_input():
    print("\n" + "=" * 60)
    ColoredOutput.print_header("邮件发送配置")
    print("=" * 60)

    smtp_server = input("SMTP 服务器地址 (如 smtp.gmail.com): ").strip() or "smtp.gmail.com"
    smtp_port_str = input("SMTP 端口 (默认 587): ").strip() or "587"
    smtp_port = int(smtp_port_str)
    use_tls = input("是否使用 TLS (y/n，默认 y): ").strip().lower() != 'n'
    sender_email = input("发件人邮箱: ").strip()
    sender_password = input("发件人邮箱授权码: ").strip()

    print("\n" + "-" * 60)
    ColoredOutput.print_info("文件路径配置")
    print("-" * 60)

    excel_path = input("学生信息 Excel 文件路径 (默认 student_info.xlsx): ").strip()
    excel_path = clean_path(excel_path) if excel_path else "student_info.xlsx"

    # 【Excel列配置说明】
    # Excel表格列号对应：A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10...
    # 请根据您的Excel表格实际列位置修改：
    # - 学号列：包含学生学号的列
    # - 姓名列：包含学生姓名的列
    # - 邮箱列：包含学生邮箱地址的列
    print("\n" + "-" * 60)
    ColoredOutput.print_info("Excel 列配置 (如 A=1, B=2, C=3)")
    print("-" * 60)

    student_id_col_str = input("学号所在列 (如 C，默认 A): ").strip().upper() or "A"
    student_id_col = ord(student_id_col_str) - ord('A') + 1

    name_col_str = input("姓名所在列 (如 B，默认 B): ").strip().upper() or "B"
    name_col = ord(name_col_str) - ord('A') + 1

    email_col_str = input("邮件地址所在列 (如 J，默认 C): ").strip().upper() or "C"
    email_col = ord(email_col_str) - ord('A') + 1

    template_path = input("邮件模板文件路径 (默认 email_template.txt): ").strip()
    template_path = clean_path(template_path) if template_path else "email_template.txt"

    send_interval_str = input("发送间隔秒数 (默认 3): ").strip()
    send_interval = int(send_interval_str) if send_interval_str else 3

    return {
        "smtp_server": smtp_server,
        "smtp_port": smtp_port,
        "use_tls": use_tls,
        "sender_email": sender_email,
        "sender_password": sender_password,
        "student_info_excel": excel_path,
        "excel_student_id_col": student_id_col,
        "excel_name_col": name_col,
        "excel_email_col": email_col,
        "email_template": template_path,
        "send_interval": send_interval
    }


def get_config_interactive():
    print("\n" + "-" * 60)
    ColoredOutput.print_info("请配置邮件发送参数")
    print("-" * 60)

    user_config = get_all_config_from_input()

    print("\n" + "-" * 60)
    print("1. 保存配置到 config.json (保存除密码外的其他配置)")
    print("2. 仅本次使用")
    choice = input("请选择 (1/2，默认 2): ").strip()

    if choice == "1":
        config = load_config()
        config.update({
            "smtp_server": user_config["smtp_server"],
            "smtp_port": user_config["smtp_port"],
            "use_tls": user_config["use_tls"],
            "sender_email": user_config["sender_email"],
            "sender_password": user_config["sender_password"],
            "student_info_excel": user_config["student_info_excel"],
            "excel_student_id_col": user_config["excel_student_id_col"],
            "excel_name_col": user_config["excel_name_col"],
            "excel_email_col": user_config["excel_email_col"],
            "email_template": user_config["email_template"],
            "send_interval": user_config["send_interval"]
        })
        save_config(config)
        ColoredOutput.print_success("配置已保存到 config.json")
    else:
        print("配置仅本次使用")

    return user_config


def load_status(status_file=STATUS_FILE):
    if os.path.exists(status_file):
        try:
            with open(status_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"警告：读取状态文件失败: {e}")
    return {
        "processed_files": [],
        "failed_files": [],
        "last_processed_file": None
    }


def save_status(status_file, status):
    try:
        with open(status_file, 'w', encoding='utf-8') as f:
            json.dump(status, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"警告：保存状态文件失败: {e}")


def cleanup_status(status_file):
    if os.path.exists(status_file):
        try:
            os.remove(status_file)
        except Exception as e:
            print(f"警告：删除状态文件失败: {e}")


def get_pending_files(all_files, status):
    processed = {item['file'] for item in status['processed_files']}
    return [f for f in all_files if f['filename'] not in processed]


def extract_student_id(filename):
    """从文件名中提取学号

    【配置说明】
    此函数用于从PDF附件文件名中提取学号，需要根据实际文件名格式修改正则表达式。

    当前支持的格式示例：
    - "签署后：2023001_001.pdf"  -> 学号: 2023001
    - "2023001_001.pdf"         -> 学号: 2023001

    【修改说明】
    如果您的文件名格式不同，请修改下面的正则表达式：
    - r'签署后：(\\d+)_'  : 匹配"签署后：学号_数字"格式
    - r'(\\d+)_'         : 匹配"学号_数字"格式

    正则表达式说明：
    - (\\d+)  : 捕获学号（连续数字）
    - _      : 下划线分隔符

    示例：如果您的文件是 "学号2023001.pdf" 格式，改为：
    match = re.search(r'学号(\\d+)', name_without_ext)
    """
    name_without_ext = os.path.splitext(filename)[0]
    match = re.search(r'签署后：(\d+)_', name_without_ext)
    if match:
        return match.group(1)
    match = re.search(r'(\d+)_', name_without_ext)
    if match:
        return match.group(1)
    return None


def get_files_with_numbers(folder_path):
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"文件夹不存在: {folder_path}")

    files = []
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        if os.path.isfile(filepath):
            student_id = extract_student_id(filename)
            if student_id:
                files.append({
                    'filename': filename,
                    'filepath': filepath,
                    'student_id': student_id
                })

    files.sort(key=lambda x: x['student_id'])
    return files


def read_student_info(excel_path, student_id, student_id_col=1, name_col=2, email_col=3):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):
            if student_id_col <= len(row):
                cell_value = row[student_id_col - 1]
                if cell_value is None:
                    continue
                cell_student_id = str(cell_value).strip()
                if '.' in cell_student_id:
                    cell_student_id = cell_student_id.split('.')[0]

                if cell_student_id == student_id:
                    name = row[name_col - 1] if name_col <= len(row) and row[name_col - 1] else ""
                    email = row[email_col - 1] if email_col <= len(row) and row[email_col - 1] else ""
                    return {
                        'name': name,
                        'email': email
                    }

        return None
    except Exception as e:
        raise Exception(f"读取 Excel 文件失败: {e}")


def export_failed_to_excel(status, folder_path):
    if not status['failed_files']:
        print("没有失败记录需要导出")
        return

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(os.path.dirname(folder_path) if folder_path else os.getcwd(),
                              f"failed_records_{timestamp}.xlsx")

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "发送失败记录"

        headers = ['序号', '文件名', '学号', '姓名', '邮件地址', '失败原因', '失败时间']
        ws.append(headers)

        for idx, item in enumerate(status['failed_files'], 1):
            ws.append([
                idx,
                item.get('file', ''),
                item.get('student_id', ''),
                item.get('name', ''),
                item.get('email', ''),
                item.get('error', ''),
                item.get('failed_time', '')
            ])

        for col in range(1, len(headers) + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20

        wb.save(output_file)
        print(f"失败记录已导出至: {output_file}")
    except Exception as e:
        print(f"导出失败: {e}")


def load_email_template(template_path):
    """加载邮件模板

    【模板格式说明】
    邮件模板文件应包含主题和正文，用空行分隔。格式如下：

    格式1（推荐）：
    Subject: 邮件主题
    <空行>
    邮件正文内容...

    格式2：
    邮件主题
    <空行>
    邮件正文内容...

    【占位符说明】
    模板中可使用以下占位符，会被自动替换：
    - "姓名" -> 学生姓名（在 replace_placeholder 函数中定义）

    【修改说明】
    如果需要添加更多占位符，请在 replace_placeholder 函数中添加相应的替换逻辑。

    【示例模板文件 email_template.txt 内容】：
    Subject: 学习资料发放通知

    姓名同学：

    您好！您的学习资料已附件形式发送，请查收。

    祝学习进步！

    管理员
    """
    template_path = str(template_path).strip()
    if template_path.startswith('"') and template_path.endswith('"'):
        template_path = template_path[1:-1]

    if not os.path.exists(template_path):
        template_path = os.path.join(os.path.dirname(__file__), os.path.basename(template_path))

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"邮件模板文件不存在: {template_path}")

    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            content = f.read()

        if '\n\n' in content:
            parts = content.split('\n\n', 1)
            subject = parts[0].replace('Subject:', '').strip()
            body = parts[1].strip()
        else:
            lines = content.split('\n')
            subject_line_idx = -1
            for i, line in enumerate(lines):
                if line.lower().startswith('subject:'):
                    subject_line_idx = i
                    break

            if subject_line_idx >= 0:
                subject = lines[subject_line_idx].replace('Subject:', '').strip()
                body = '\n'.join(lines[subject_line_idx + 1:]).strip()
            else:
                subject = "学习资料"
                body = content

        return {'subject': subject, 'body': body}
    except Exception as e:
        raise Exception(f"读取邮件模板失败: {e}")


def replace_placeholder(text, student_name):
    """替换邮件模板中的占位符

    【配置说明】
    此函数用于将邮件模板中的占位符替换为实际的学生信息。

    当前支持的占位符：
    - "姓名" -> 学生姓名

    【修改说明】
    如果您在邮件模板中使用了其他占位符，需要在这里添加相应的替换逻辑。

    示例：
    如果模板中有 "学号" 占位符，添加：
    text = text.replace("学号", student_id)

    如果需要多个替换，可以链式调用：
    text = text.replace("占位符1", value1).replace("占位符2", value2)
    """
    return text.replace("姓名", student_name)


def send_email(smtp_config, recipient, subject, body, attachment_path=None):
    """发送邮件

    【配置说明】
    此函数负责邮件的实际发送。

    【附件说明】
    - attachment_path: PDF附件的完整文件路径
    - 程序会自动将文件作为附件添加到邮件中
    - 附件名称使用原始文件名（通过 os.path.basename 获取）

    【SMTP配置说明】
    smtp_config 字典应包含：
    - smtp_server: SMTP服务器地址
    - smtp_port: SMTP端口
    - use_tls: 是否使用TLS加密
    - sender_email: 发件人邮箱
    - sender_password: 发件人邮箱授权码（不是邮箱密码）

    【授权码获取】
    大多数邮箱需要使用授权码而非登录密码：
    - QQ邮箱：设置 -> 账户 -> 开启SMTP服务 -> 获取授权码
    - 163邮箱：设置 -> POP3/SMTP -> 开启 -> 获取授权码
    - Gmail：需要开启"低安全性应用访问"或使用应用专用密码
    """
    msg = MIMEMultipart()
    msg['From'] = smtp_config['sender_email']
    msg['To'] = recipient
    msg['Subject'] = Header(subject, 'utf-8').encode()
    msg.attach(MIMEText(body, 'plain', 'utf8'))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            attachment_data = f.read()
        filename = os.path.basename(attachment_path)
        part = MIMEApplication(attachment_data, Name=filename)
        part.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', filename))
        msg.attach(part)

    server = None
    try:
        port = smtp_config.get('smtp_port', 587)
        use_ssl = port in [465, 994] or smtp_config.get('use_ssl', False)
        timeout = 30

        server = smtplib.SMTP(smtp_config['smtp_server'], port, timeout=timeout)

        if use_ssl:
            server.starttls()
        elif smtp_config.get('use_tls', False):
            server.starttls()

        server.login(smtp_config['sender_email'], smtp_config['sender_password'])

        if '@' not in recipient:
            raise Exception(f"收件人地址无效: {recipient}")

        server.sendmail(smtp_config['sender_email'], [recipient], msg.as_string())
        return True
    except smtplib.SMTPAuthenticationError:
        raise Exception("SMTP 认证失败，请检查邮箱和密码/授权码")
    except smtplib.SMTPSenderRefused as e:
        raise Exception(f"发件人被拒绝: {e}，发件人邮箱地址无效或不存在")
    except smtplib.SMTPRecipientsRefused as e:
        raise Exception(f"收件人被拒绝: {e}，收件人邮箱地址无效或不存在")
    except smtplib.SMTPConnectError as e:
        raise Exception(f"无法连接到 SMTP 服务器 {smtp_config['smtp_server']}:{port}，请检查服务器地址、端口和网络连接")
    except smtplib.SMTPServerDisconnected:
        raise Exception("SMTP 服务器断开连接，可能是端口或 TLS 设置不正确")
    except TimeoutError:
        raise Exception(f"连接 SMTP 服务器超时，请检查网络或防火墙设置")
    except OSError as e:
        errno = getattr(e, 'errno', 0)
        if errno == 11001:
            raise Exception(f"无法解析域名 {smtp_config['smtp_server']}，请检查服务器地址是否正确")
        elif errno == 10060:
            raise Exception(f"连接超时，请检查网络或防火墙设置")
        elif errno == 10061:
            raise Exception(f"服务器拒绝连接，请检查端口是否正确")
        else:
            raise Exception(f"网络错误 (代码 {errno}): {e}")
    except Exception as e:
        raise Exception(f"邮件发送失败: {str(e)}")
    finally:
        if server:
            try:
                server.quit()
            except:
                pass


def send_email_with_retry(smtp_config, recipient, subject, body, attachment_path, max_retries, retry_interval):
    for attempt in range(max_retries):
        try:
            send_email(smtp_config, recipient, subject, body, attachment_path)
            return True, None
        except Exception as e:
            if attempt < max_retries - 1:
                error_msg = f"网络错误，{retry_interval}秒后重试... ({attempt + 1}/{max_retries})"
                ColoredOutput.print_warning(error_msg)
                time.sleep(retry_interval)
            else:
                return False, str(e)
    return False, "未知错误"


def parse_arguments():
    parser = argparse.ArgumentParser(description='批量发送邮件程序')
    parser.add_argument('-p', '--path', type=str, help='指定附件文件夹路径')
    parser.add_argument('-r', '--restart', action='store_true', help='强制重新开始，清除之前的状态')
    parser.add_argument('-v', '--verbose', action='store_true', help='显示详细日志')
    parser.add_argument('-s', '--status', action='store_true', help='查看当前状态（不发送邮件）')
    parser.add_argument('--smtp', type=str, help='SMTP 服务器地址')
    parser.add_argument('--port', type=int, help='SMTP 端口')
    parser.add_argument('--email', type=str, help='发件人邮箱')
    parser.add_argument('--password', type=str, help='发件人密码/授权码')
    parser.add_argument('--config', type=str, help='指定配置文件路径')
    return parser.parse_args()


def get_folder_path(args):
    folder_path = args.path
    while not folder_path:
        folder_path = clean_path(input("\n请输入附件文件夹路径: ").strip())
    return folder_path


def create_empty_status():
    return {"processed_files": [], "failed_files": [], "last_processed_file": None}


def get_retry_files(status, folder_path):
    if not status['failed_files']:
        return []
    
    retry_files = []
    for failed in status['failed_files']:
        filename = failed.get('file')
        if not filename:
            continue
        filepath = os.path.join(folder_path, filename)
        if os.path.exists(filepath):
            student_id = failed.get('student_id', '')
            if not student_id:
                student_id = extract_student_id(filename) or ''
            retry_files.append({
                'filename': filename,
                'filepath': filepath,
                'student_id': student_id
            })
    
    retry_files.sort(key=lambda x: x['student_id'])
    return retry_files


def handle_resume_option(status, files, folder_path):
    pending_files = get_pending_files(files, status)

    if len(status['processed_files']) > 0 or len(status['failed_files']) > 0:
        print("\n" + "=" * 60)
        ColoredOutput.print_warning("检测到之前的发送进度!")
        print("=" * 60)
        print(f"  已成功发送: {len(status['processed_files'])} 封")
        print(f"  失败记录: {len(status['failed_files'])} 条")
        print(f"  剩余待发送: {len(pending_files)} 封")
        if status.get('last_processed_file'):
            print(f"  上次中断于: {status['last_processed_file']}")
        print("=" * 60)
        print()
        print("  1. 从断点继续 (推荐)")
        print("  2. 重新开始")
        print("  3. 仅导出失败记录到 Excel")
        if status['failed_files']:
            print("  4. 重试失败的文件")
            print("  5. 退出程序")
        else:
            print("  4. 退出程序")
        print()
        
        if status['failed_files']:
            choice = input("请选择 (1/2/3/4/5，默认 1): ").strip().lower()
        else:
            choice = input("请选择 (1/2/3/4，默认 1): ").strip().lower()
        print()

        if choice == '2':
            cleanup_status(STATUS_FILE)
            print("已清除进度，重新开始")
            return files
        elif choice == '3':
            export_failed_to_excel(status, folder_path)
            print("已导出失败记录，程序退出")
            return None
        elif choice == '4' and status['failed_files']:
            retry_files = get_retry_files(status, folder_path)
            if retry_files:
                print(f"将重试 {len(retry_files)} 个失败的文件")
                return retry_files
            else:
                print("没有可重试的文件")
                return None
        elif (choice == '5' and status['failed_files']) or (choice == '4' and not status['failed_files']):
            print("程序退出")
            return None
        else:
            print("从断点继续...")
            return pending_files

    return pending_files


def ensure_config_complete(config, args):
    if args.smtp:
        config['smtp_server'] = args.smtp
    if args.port:
        config['smtp_port'] = args.port
    if args.email:
        config['sender_email'] = args.email
    if args.password:
        config['sender_password'] = args.password

    if config.get('sender_email') and config.get('sender_password'):
        ColoredOutput.print_info(f"已加载配置: {config.get('sender_email')}")
        choice = input("是否使用此配置? (Y/n，默认 Y): ").strip().lower()
        if choice == 'n':
            user_config = get_config_interactive()
            config.update(user_config)
        else:
            print("使用现有配置")
    else:
        user_config = get_config_interactive()
        config.update(user_config)

    return config


def build_smtp_config(config):
    return {
        'smtp_server': config['smtp_server'],
        'smtp_port': config['smtp_port'],
        'use_tls': config.get('use_tls', False),
        'sender_email': config['sender_email'],
        'sender_password': config['sender_password']
    }


def show_status(status):
    print("\n" + "=" * 60)
    ColoredOutput.print_header("当前处理状态")
    print("=" * 60)
    print(f"已成功发送: {len(status['processed_files'])} 封")
    print(f"失败: {len(status['failed_files'])} 封")

    if status['failed_files']:
        print("\n失败文件列表:")
        for item in status['failed_files']:
            print(f"  - {item['file']}: {item['error']}")

    print("=" * 60)


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir:
        os.chdir(script_dir)

    print("\n" + "=" * 60)
    ColoredOutput.print_header("批量发送邮件程序")
    print("=" * 60)

    args = parse_arguments()

    if args.status:
        status = load_status()
        show_status(status)
        return

    config = load_config()
    config = ensure_config_complete(config, args)

    folder_path = get_folder_path(args)

    try:
        files = get_files_with_numbers(folder_path)
    except FileNotFoundError as e:
        ColoredOutput.print_error(str(e))
        return

    if not files:
        ColoredOutput.print_warning("未找到符合条件的文件（文件名需包含下划线和数字）")
        return

    print(f"找到 {len(files)} 个文件")

    status = load_status()

    if args.restart:
        cleanup_status()
        status = create_empty_status()
        print("已清除之前的进度")

    pending_files = handle_resume_option(status, files, folder_path)

    if pending_files is None:
        return

    smtp_config = build_smtp_config(config)

    print("\n" + "-" * 60)
    print(f"待发送文件数: {len(pending_files)}")
    send_interval = config.get('send_interval', 3)
    print(f"预计耗时: 约 {len(pending_files) * send_interval} 秒")
    print("-" * 60)

    confirm = input("\n确认开始发送? (Y/n，默认 Y): ").strip().lower()
    if confirm == 'n':
        print("程序退出")
        return

    print()

    try:
        template = load_email_template(config.get('email_template', 'email_template.txt'))
        excel_path = config.get('student_info_excel', 'student_info.xlsx')

        progress_display = PowerShellProgress(len(pending_files), send_interval)

        for file_info in pending_files:
            filename = file_info['filename']
            filepath = file_info['filepath']
            student_id = file_info['student_id']

            student_id_col = config.get('excel_student_id_col', 1)
            name_col = config.get('excel_name_col', 2)
            email_col = config.get('excel_email_col', 3)

            try:
                student_info = read_student_info(excel_path, student_id, student_id_col, name_col, email_col)

                if not student_info:
                    error_msg = f"学号 {student_id} 在 Excel 中未找到"
                    status['failed_files'].append({
                        'file': filename,
                        'student_id': student_id,
                        'error': error_msg,
                        'failed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    save_status(STATUS_FILE, status)
                    progress_display.update(filename, "未知", "未知", "failed", error_msg)
                    time.sleep(send_interval)
                    continue

                name = student_info['name']
                email = student_info['email']

                if not email:
                    error_msg = "邮件地址为空"
                    status['failed_files'].append({
                        'file': filename,
                        'student_id': student_id,
                        'name': name,
                        'error': error_msg,
                        'failed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    save_status(STATUS_FILE, status)
                    progress_display.update(filename, name, email, "failed", error_msg)
                    time.sleep(send_interval)
                    continue

                subject = replace_placeholder(template['subject'], name)
                body = replace_placeholder(template['body'], name)

                max_retries = config.get('max_retries', 3)
                retry_interval = config.get('retry_interval', 10)
                success, error_msg = send_email_with_retry(
                    smtp_config, email, subject, body, filepath,
                    max_retries, retry_interval
                )

                if success:
                    status['processed_files'].append({
                        'file': filename,
                        'student_id': student_id,
                        'name': name,
                        'email': email,
                        'status': 'success',
                        'sent_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    progress_display.update(filename, name, email, "success")
                else:
                    status['failed_files'].append({
                        'file': filename,
                        'student_id': student_id,
                        'name': name,
                        'email': email,
                        'error': error_msg,
                        'failed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    progress_display.update(filename, name, email, "failed", error_msg)

                status['last_processed_file'] = filename
                save_status(STATUS_FILE, status)

            except FileNotFoundError as e:
                error_msg = str(e)
                status['failed_files'].append({
                    'file': filename,
                    'student_id': student_id,
                    'error': error_msg,
                    'failed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
                save_status(STATUS_FILE, status)
                progress_display.update(filename, "未知", "未知", "failed", error_msg)

            except Exception as e:
                error_msg = f"处理失败: {str(e)}"
                status['failed_files'].append({
                    'file': filename,
                    'student_id': student_id,
                    'error': error_msg,
                    'failed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
                save_status(STATUS_FILE, status)
                progress_display.update(filename, "未知", "未知", "failed", error_msg)

            if pending_files.index(file_info) < len(pending_files) - 1:
                time.sleep(send_interval)

        progress_display.print_summary()

        if status['failed_files']:
            print("\n失败文件已记录")
            export_failed_to_excel(status, folder_path)

        if len(status['processed_files']) > 0:
            print("\n发送完成!")
            print(f"成功: {len(status['processed_files'])} 封")
            print(f"失败: {len(status['failed_files'])} 封")

    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
        ColoredOutput.print_info("进度已保存，重新运行程序将从断点继续")
        save_status(STATUS_FILE, status)
    except FileNotFoundError as e:
        ColoredOutput.print_error(str(e))
    except Exception as e:
        ColoredOutput.print_error(f"发生错误: {str(e)}")


if __name__ == "__main__":
    main()