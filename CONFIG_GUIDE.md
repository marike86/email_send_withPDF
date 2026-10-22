# 配置指南

桌面版默认使用 `%LOCALAPPDATA%\PdfMailer\config.json`。也可以通过 `--config` 指定，或在界面导入 `config.example.json` 的副本。

| 字段 | 说明 |
|---|---|
| smtp_server | 邮箱提供商的 SMTP 主机名 |
| smtp_port | 1–65535；465 自动使用隐式 SSL |
| use_ssl / use_tls | 隐式 SSL / STARTTLS，SSL 优先，发送必须加密 |
| sender_email | 单个发件邮箱 |
| student_info_excel | `.xlsx` 学生表，首行为表头，工作表可通过导入向导选择 |
| email_template | UTF-8 模板，第一行为主题，其余为正文 |
| attachment_folder | PDF 附件目录 |
| state_file | 原生桌面版发送日志；CLI 请显式使用 --state |
| excel_student_id_col | 学号列，默认 1（A） |
| excel_name_col | 姓名列，默认 2（B） |
| excel_email_col | 邮箱列，默认 3（C） |
| excel_worksheet | 原生桌面版工作表名称，空值使用活动工作表 |
| group_attachments | 原生桌面版按学号合并附件，默认 false |
| send_interval | 每封邮件间隔秒数，默认 3 |
| max_retries | 含首次在内的最多尝试次数，1–10，默认 3 |
| retry_interval | 临时故障重试间隔，默认 10 秒 |

间隔允许小数，原生版取 0–86400 秒。列号支持整数和 A、AA 等字母，三列必须不同。学号必须唯一；前导零应实际存为文本，仅设置 Excel 显示格式无法恢复。

新增的工作表与多附件选项仅用于原生桌面版；Python CLI 保留原行为。

配置中的相对路径以配置文件目录为基准。导入后会转换为完整路径。环境变量使用对应大写名，如 `SMTP_SERVER`、`SMTP_PORT`、`USE_SSL`、`SENDER_EMAIL`、`SENDER_PASSWORD`；布尔值为 true/false 或 1/0。环境变量在加载配置时覆盖文件内容。

授权码在界面中输入，默认仅本次使用。主动勾选“安全记住授权码”并保存账户或设置后，由 Windows 凭据保管库保存。取消勾选后保存或删除账户，会移除保存的凭据。配置保存始终排除 `sender_password`。可读取旧配置中的密码兼容迁移，保存后移除。

## 模板

```text
Subject: {姓名}同学的资料

{姓名}同学：
您好，学号 {学号} 的资料请查收附件。
```

主题也可省略 `Subject:` 前缀。兼容旧模板中裸写的“姓名”，因此普通正文不宜包含其他用途的“姓名”字样。原生版单次替换，不会再将插入的姓名内容当作新占位符。

文件名支持 `学号_说明.pdf`、`学号.pdf`、`签署后：学号_说明.pdf` 和半角冒号形式，学号为连续数字。扩展名不区分大小写；不从任意文件名中截取一段数字。

## SMTP

465 使用 SSL/TLS；其他端口配合 STARTTLS。自定义隐式 SSL 端口勾选 SSL/TLS。MailKit 默认验证证书，不关闭验证、不明文认证。请按邮箱提供商的当前说明获取授权码或应用专用密码。

连接故障及 SMTP 临时响应可重试；认证失败、附件错误等不会反复重试。提交时断线无法确认是否送达，将标为待核实。

## 日志迁移

原生桌面版与上一版 Python 的 `version: 2` 日志兼容，同一附件实际路径、内容、账户、收件人、主题及正文会生成相同身份。导入配置时会查找配置旁边的 `status.json`，请核对路径。

最早版日志仅按文件名记录，无法可靠判断发送内容，新版会报错并保留文件。请备份并核对邮箱后，使用新的日志或将确认未发送的附件放到单独目录。不自动删除历史，不静默重置。
