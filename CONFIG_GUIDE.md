# 配置指南

复制 `config.example.json` 为 `config.json`，或在桌面版保存设置。

| 配置项 | 说明 |
|---|---|
| smtp_server | 邮箱提供商的 SMTP 主机名 |
| smtp_port | 1–65535；465 自动使用隐式 SSL |
| use_ssl / use_tls | 隐式 SSL / STARTTLS；发信必须加密；use_ssl 优先 |
| sender_email | 单个发件邮箱地址 |
| student_info_excel | 学生信息 `.xlsx`，第一行为表头，读取活动工作表 |
| email_template | UTF-8 文本模板；第一行为主题，其余为正文 |
| attachment_folder | PDF 附件目录，也可通过 `-p` 指定 |
| excel_student_id_col | 学号列，默认 1（A），支持 AA 等字母 |
| excel_name_col | 姓名列，默认 2（B） |
| excel_email_col | 邮箱列，默认 3（C） |
| send_interval | 邮件间隔秒数，默认 3，可为非负小数 |
| max_retries | 最多尝试次数，含首次，1–10，默认 3 |
| retry_interval | 临时故障重试间隔秒数，默认 10 |

所有配置项都支持对应大写环境变量，例如 `SMTP_SERVER`、`SMTP_PORT`、`USE_TLS`、`USE_SSL`、`SENDER_EMAIL`、`SENDER_PASSWORD`。布尔变量支持 true/false 或 1/0。环境变量优先于配置文件，命令行 `--smtp`、`--port`、`--email` 再覆盖对应参数。

授权码应在界面或命令行隐藏输入，或通过 `SENDER_PASSWORD` 提供。保存配置始终排除 `sender_password`。可保留已有配置中的密码供旧版迁移，但建议保存一次以移除。

配置内的相对路径以配置文件目录为基准；命令行路径相对当前工作目录。模板找不到时会报错，不再擅自替换为同名文件。

## 模板和文件名

```text
Subject: {姓名}同学的资料

{姓名}同学：
您好，学号 {学号} 的资料请查收附件。
```

也支持第一行没有 `Subject:` 前缀。兼容旧模板中裸写的“姓名”，因此普通正文不宜包含其他用途的“姓名”字样。

文件名支持 `学号_说明.pdf`、`学号.pdf`、`签署后：学号_说明.pdf` 和半角冒号形式。学号为连续数字。为避免匹配到错误的人，不再从任意文件名中截取一段数字。

Excel 学号必须唯一；缺失姓名或无效邮箱会阻止对应附件发送；重复学号会阻止整个预检。文本前导零会保留；Excel 数字格式显示出来的前导零无法恢复，请将学号实际保存为文本。

## 加密与错误处理

465 使用 `SMTP_SSL`；其他端口配合 `use_tls: true` 使用 STARTTLS；自定义隐式 SSL 端口设置 `use_ssl: true`。证书验证默认开启。请以邮箱提供商公布的 SMTP 配置和认证方式为准。

仅连接故障和 SMTP 临时响应会自动重试；认证失败、永久拒绝、附件问题不会自动反复发送。提交阶段断线无法确定是否送达，状态为“待核实”。在桌面版核对邮箱后处理，或保留日志并人工核实。

发送日志格式、旧日志迁移及命令行示例见 [README](README.md)。
