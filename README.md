# PDF 邮件工作台

根据 PDF 文件名中的学号匹配 Excel 学生信息，预览后批量发送邮件。

桌面版采用 **WinUI 3 / Fluent 视觉风格**：浅色卡片、侧边导航、圆角输入和蓝色主操作，使用 Python + Qt 实现；不是原生 Windows App SDK 应用。

## 快速开始

需要 Python 3.10 或更高版本。桌面版安装：

```powershell
python -m pip install -r requirements-gui.txt
python send_email_gui.py
```

仅使用命令行时安装 `requirements.txt`，无需 Qt。

1. 准备 `.xlsx`：第一行表头，默认 A=学号、B=姓名、C=邮箱；学号含前导零时存为文本。
2. 准备 PDF 目录：支持 `2023001_资料.pdf`、`签署后：2023001_001.pdf`、`2023001.pdf`，扩展名不区分大小写。
3. 选择 UTF-8 模板，可直接使用 [示例模板](examples/email_template.txt)。
4. 在“账户与设置”填写发件邮箱、SMTP 服务器、端口、加密方式和授权码。
5. 点击“检查文件”，逐行预览收件人、附件、主题和正文，修正需关注项。
6. 点击“开始发送”，核对发送数量并确认。需修正和待核实项不会发送。

授权码不会写入配置文件。可点击“保存设置”保存其他参数；下次会话需重新输入授权码，或通过 `SENDER_PASSWORD` 环境变量提供。

## 功能

- 离线预检：重复学号、无效邮箱、未匹配学号、非 PDF 内容、缺失姓名均明确提示。
- 一次读取 Excel，支持 A、AA 等多字母列，及时释放文件句柄。
- 逐封邮件预览，按姓名、邮箱、学号、文件名搜索，按状态筛选。
- 后台发送不阻塞界面，支持停止后续邮件；当前 SMTP 请求完成或超时后停止。
- SSL（通常 465）与 STARTTLS（通常 587）均校验证书，拒绝明文认证。
- 原子写入发送日志，自动跳过已成功发送的同一任务；失败重试成功后覆盖旧失败记录。
- 提交中断或发送时崩溃的邮件标记“待核实”，不自动重发。核对邮箱后可标记成功或允许重试。
- 导出全部预检和发送结果为 XLSX 或 Excel 可读的 UTF-8 BOM CSV。

“待发送”统计包含可重试失败项；“需关注”包含失败、需修正及待核实记录。

## 命令行

复制 `config.example.json` 为 `config.json` 并修改路径和账户。

```powershell
# 只检查，不连接 SMTP，不要求授权码
python send_email.py --config config.json -p "D:\PDF" --dry-run --export results.csv

# 预检后输入 SEND 才发送；授权码隐藏输入
python send_email.py --config config.json -p "D:\PDF"

# 只重试日志中的失败项
python send_email.py --config config.json -p "D:\PDF" --retry-failed

# 查看日志
python send_email.py --status

# 分别指定配置和发送日志，GUI 也支持这两个选项
python send_email_gui.py --config "D:\任务\config.json" --state "D:\任务\status.json"
```

配置路径相对当前工作目录解析；配置中相对文件路径以配置文件所在目录为基准。`--state`、`--path` 和 `--export` 相对当前工作目录解析，程序不再切换工作目录。CLI 退出码：0=无错误，1=预检/发送错误或待核实，130=用户中断。

## 发送历史与兼容性

日志是 `version: 2` 的 JSON，以附件绝对路径、内容指纹、发件人、收件人、主题及正文共同识别任务。移动附件、更换账户/收件人或修改内容会形成新任务，发送前务必核对计划。多个进程共用同一日志时有系统文件锁；不要使用不同日志同时发送同一批文件。

旧版 `status.json` 仅按文件名记录，无法可靠判断账户和内容。新版遇到旧日志或损坏日志会停止，不会静默重置。请备份旧日志、核对已发送邮件后，用 `--state` 指定新的日志，只向确认尚未送达的附件发送。`--restart` 不再删除历史；`--password` 参数已移除，改用隐藏输入或环境变量。

已有旧配置中的授权码仍可读取以兼容旧版本；保存设置会从配置中移除它。建议迁移后保存一次配置。

SMTP 接受提交不等于收件人最终收到邮件，退信仍需在邮箱检查。PDF 标识检查仅验证文件开头，不进行完整 PDF 内容校验。

## 验证

```powershell
python -m unittest -v test_send_email test_gui
```

测试覆盖发送协议、续发、重试、异常断线、日志写入失败、停止、配置、预检、CSV 及 GUI 后台交互。SMTP 全部模拟，不发送真实邮件。GUI 测试使用离屏模式；未安装 Qt 时会跳过 GUI 测试。CI 配置覆盖 Windows/Linux 与 Python 3.10/3.14。

更多配置见 [配置指南](CONFIG_GUIDE.md)，操作说明见 [使用说明](使用说明.md)。
