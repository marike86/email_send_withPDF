# 配置指南

本文档说明 `send_email.py` 中需要用户根据实际情况修改的配置项。

---

## 1. 文件名中学号提取

**位置**: `extract_student_id()` 函数（约第219行）

**当前支持的格式**:
- `签署后：2023001_001.pdf` → 学号: `2023001`
- `2023001_001.pdf` → 学号: `2023001`

**修改方法**: 如果您的文件名格式不同，请修改正则表达式：

```python
# 当前代码
match = re.search(r'签署后：(\d+)_', name_without_ext)
match = re.search(r'(\d+)_', name_without_ext)

# 示例：如果是 "学号2023001.pdf" 格式
match = re.search(r'学号(\d+)', name_without_ext)

# 示例：如果是 "2023001-附件.pdf" 格式
match = re.search(r'(\d+)-', name_without_ext)
```

**正则表达式说明**:
| 符号 | 含义 |
|------|------|
| `\d+` | 匹配一个或多个数字 |
| `(\d+)` | 捕获数字到第一组 |
| `_` | 匹配下划线 |
| `-` | 匹配连字符 |
| `:` | 匹配冒号 |

---

## 2. 邮件模板占位符

**位置**: `replace_placeholder()` 函数（约第382行）

**当前支持的占位符**:
- `姓名` → 被替换为学生姓名

**修改方法**: 如需添加更多占位符：

```python
def replace_placeholder(text, student_name, student_id=None):
    text = text.replace("姓名", student_name)
    if student_id:
        text = text.replace("学号", student_id)  # 添加学号占位符
    return text
```

**模板文件格式** (`email_template.txt`):
```
Subject: 邮件主题

姓名同学：

您好！邮件正文...

管理员
```

---

## 3. Excel 学生信息表

**位置**: 交互式配置（约第120-140行）

**列号对应表**:
| 列字母 | 列号 | 说明 |
|--------|------|------|
| A      | 1    | 第1列 |
| B      | 2    | 第2列 |
| C      | 3    | 第3列 |
| D      | 4    | 第4列 |
| E      | 5    | 第5列 |
| F      | 6    | 第6列 |
| G      | 7    | 第7列 |
| H      | 8    | 第8列 |
| I      | 9    | 第9列 |
| J      | 10   | 第10列 |

**需要配置的列**:
- 学号所在列（如：A列 = 1）
- 姓名所在列（如：B列 = 2）
- 邮箱所在列（如：J列 = 10）

---

## 4. SMTP 邮件服务器

**位置**: `DEFAULT_CONFIG`（第28行）或交互式配置

### 常用邮箱配置

| 邮箱服务 | SMTP服务器 | 端口 | TLS | 说明 |
|----------|------------|------|-----|------|
| Gmail | smtp.gmail.com | 587 | 是 | 需要开启"低安全性应用访问" |
| QQ邮箱 | smtp.qq.com | 587 | 是 | 需要获取授权码 |
| 163邮箱 | smtp.163.com | 587 | 是 | 需要获取授权码 |
| 企业邮箱 | smtp.您的域名.com | 25/587 | 根据服务器 | 联系管理员获取 |

### 端口说明

| 端口 | 加密方式 | 说明 |
|------|----------|------|
| 25 | 无/TLS | 最常用，非加密或STARTTLS |
| 465 | SSL | 隐式SSL加密 |
| 587 | TLS | 显式TLS加密（推荐）|

### 授权码获取

- **QQ邮箱**：设置 → 账户 → 开启SMTP服务 → 获取授权码
- **163邮箱**：设置 → POP3/SMTP → 开启 → 获取授权码
- **Gmail**：Google账户 → 安全 → 开启"低安全性应用访问"

---

## 5. 发送配置

**位置**: `DEFAULT_CONFIG`（第28行）

```python
"send_interval": 3,      # 发送间隔秒数
"max_retries": 3,        # 失败重试次数
"retry_interval": 10   # 重试间隔秒数
```

### 配置建议

| 配置项 | 推荐值 | 说明 |
|--------|--------|------|
| send_interval | 3-5秒 | 避免被识别为垃圾邮件 |
| max_retries | 3次 | 网络不稳定时可增加 |
| retry_interval | 10秒 | 给服务器时间恢复 |

---

## 6. 默认配置

**位置**: `DEFAULT_CONFIG` 字典（第28行）

```python
DEFAULT_CONFIG = {
    # SMTP服务器配置
    "smtp_server": "smtp.gmail.com",  # 修改为您的SMTP服务器
    "smtp_port": 587,                  # 修改为您的端口
    "use_tls": True,                   # 根据服务器设置

    # 文件路径配置
    "email_template": "email_template.txt",  # 模板文件路径
    "student_info_excel": "student_info.xlsx", # Excel文件路径

    # 发送配置
    "send_interval": 3,       # 间隔秒数
    "max_retries": 3,         # 重试次数
    "retry_interval": 10      # 重试间隔
}
```

---

## 7. 配置文件

程序会创建 `config.json` 保存配置，格式如下：

```json
{
  "smtp_server": "smtp.qq.com",
  "smtp_port": 587,
  "use_tls": true,
  "sender_email": "your_email@qq.com",
  "sender_password": "your_auth_code",
  "email_template": "email_template.txt",
  "student_info_excel": "student_info.xlsx",
  "excel_student_id_col": 1,
  "excel_name_col": 2,
  "excel_email_col": 10,
  "send_interval": 3
}
```

---

## 8. 常见问题

### Q: 提示"学号在Excel中未找到"
A:
- 检查Excel中学号列是否正确
- 检查文件名中学号与Excel中是否完全一致（包括数字类型，如 2023001 vs 2023001.0）
- 确认学号没有多余空格

### Q: 提示"SMTP认证失败"
A:
- 确认使用的是**授权码**而非登录密码
- 部分邮箱需要使用应用专用密码

### Q: 提示"收件人邮箱地址无效"
A:
- 检查Excel中邮箱地址格式是否正确
- 确认邮箱域名是否存在

### Q: 如何只发送失败的文件？
A: 运行程序后选择菜单中的"重试失败的文件"选项

### Q: 如何导出失败记录？
A: 运行程序后选择菜单中的"仅导出失败记录到 Excel"选项

---

## 快速开始

1. 准备Excel表格（学号、姓名、邮箱）
2. 准备邮件模板
3. 将PDF文件放入文件夹，文件名包含学号
4. 运行程序，按提示配置
5. 确认发送
