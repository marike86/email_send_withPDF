# PDF 邮件工作台 · 原生 WinUI 3

按 PDF 文件名中的学号匹配 Excel 学生信息，预览后批量发送。桌面应用已迁移到 **C# + .NET 10 + Windows App SDK / WinUI 3 + XAML**，不再使用 Qt 或 Python 桌面界面。

## 使用桌面版

从主分支 Actions 的 Windows 构建产物获取 `PdfMailer-win-x64`，完整解压后运行 `PdfMailer.exe`。自包含目录中带有 .NET 和 Windows App SDK 运行时，不需要安装 Python。请保留目录内全部依赖文件。

目标系统为 64 位 Windows 11 / Windows 10 2004 及以上。桌面版采用原生 NavigationView、TextBox、ComboBox、NumberBox、文件选择器和 ContentDialog，支持系统浅色/深色主题及 DPI 缩放。

1. 准备 `.xlsx`：第一行表头，包含唯一学号、姓名和单个邮箱；含前导零的学号存为文本。
2. 准备 PDF：如 `2023001_资料.pdf`、`签署后：2023001_001.pdf`、`2023001.pdf`。
3. 选择 PDF 目录、Excel 和 UTF-8 邮件模板，可使用 `examples/email_template.txt`。
4. 在“账户与设置”填写邮箱、授权码、SMTP、加密方式及 Excel 列映射。
5. 点击“检查文件”，逐行查看匹配结果和邮件预览。
6. 点击“开始发送”，核对发件账户和数量后确认。

邮件预览使用单层原生 ScrollViewer 和可选中文本，自动换行、纵向滚动，无嵌套文本编辑框和横向滚动条。切换邮件会回到正文顶部。发送范围可选择“仅勾选项”（默认）、“当前筛选结果”或“整份计划”；确认窗口显示实际收件地址数、邮件数与附件数。重试也遵循所选范围。

## 新增工作流

- **Excel 导入向导**：选择工作表，按表头识别学号/姓名/邮箱，预览前五行并调整列映射。工作表不存在时要求重新导入。
- **范围发送**：逐项勾选、勾选或清除筛选结果；隐藏的勾选项仍属于“仅勾选项”，只发可见结果时请选择“当前筛选结果”。每次重新预检默认勾选全部记录。
- **连接检测、验证登录与试发**：设置页分别检查加密连接和认证；工作台将选中邮件试发到发件人自己的地址，主题加 `[试发]`，正文和附件与样本一致。试发只尝试一次，使用独立 `test-status.json`，不改变正式发送记录。SMTP 接受后仍需到自己的邮箱核对；不确定时先查邮箱，再决定是否重新试发。
- **PDF 原生预览**：选择附件，翻页与缩放；损坏、加密或已移动的文件会显示错误，不影响正文查看。
- **模板库**：内置主题/正文编辑，插入 `{姓名}`、`{学号}` 并实时预览。支持新建、更新、保存和应用模板，应用后重新预检。
- **任务历史**：手动命名保存，发送前后自动更新结果摘要。恢复和复制均保留原日志路径，重新检查后跳过已发送项；历史保存的是配置与摘要，详细发送记录仍在原日志中。
- **多附件合并**：开启后，同一学号的 PDF 合成一封邮件。成功或待核实状态同步到单附件身份，切换模式不会自动重发。已有部分成功或待核实的组会阻止合并，需关闭合并后核对处理。合并默认关闭。
- **多账户与安全凭据**：保存/切换/删除发件账户。主动勾选“安全记住授权码”后，由 Windows 凭据保管库保存；配置、任务历史与账户 JSON 均不包含授权码。取消勾选并保存账户/设置，或删除账户，会移除保存的凭据。

## 功能与可靠性

- 离线预检：拒绝重复学号，提示无匹配信息、缺失姓名、无效邮箱和非 PDF 内容。
- Excel 一次加载，支持 A、AA 等多字母列，及时释放文件。
- MailKit 发送 SSL/STARTTLS 邮件并验证证书，授权码不保存到配置。
- 后台任务、停止后续邮件、失败重试和断点续发。
- 原子日志、跨进程文件锁、提交前检查点，日志写入失败停止发送。
- 提交阶段断线或中断标记“待核实”，不自动重发；核实后可标记成功或允许重试。
- XLSX/UTF-8 BOM CSV 导出，保留文本学号，避免内容被识别为公式。

SMTP 接受提交不等于最终投递成功，退信需在邮箱核对。PDF 检查仅验证文件开头标识，不解析整份文件。

## 配置与迁移

默认配置和日志位于 `%LOCALAPPDATA%\PdfMailer\config.json`、`status.json`。保存设置不写入授权码；主动开启安全保存时使用 Windows 凭据保管库，也可用 `SENDER_PASSWORD` 环境变量提供。

“账户与设置 → 导入已有配置”可读取原 Python 配置，配置中的相对路径以原配置文件目录为基准。若旁边存在 `status.json`，优先使用该日志；随后保存设置会把非敏感参数写到当前应用配置。旧配置中的明文授权码只兼容读取，保存时会移除。

桌面版兼容上一版 `version: 2` 日志的身份哈希；附件绝对路径、内容、账户、收件人、主题和正文相同时可继续跳过成功项。Windows 路径会规范化大小写和实际路径。移动附件或改内容会产生新任务，发送前务必核对计划。

更早的 `processed_files` 格式及损坏日志会明确报错，不静默重置。请备份、核对已发送邮件后再改用新日志，只处理确认未送达的文件。同批任务应始终使用同一日志；跨语言同时使用时也必须指定相同日志。

```powershell
PdfMailer.exe --config "D:\邮件任务\config.json" --state "D:\邮件任务\status.json"
# 演示数据，不连接 SMTP、不保存配置
PdfMailer.exe --demo
```

## 从源码构建

需要 .NET 10 SDK 和 Windows 构建环境。Windows SDK 工具和 Windows App SDK 通过已固定版本的 NuGet 包还原。

```powershell
dotnet test native/PdfMailer.Tests/PdfMailer.Tests.csproj -c Release
dotnet build native/PdfMailer.App/PdfMailer.App.csproj -c Release -p:Platform=x64
# 验证、发布自包含桌面版，并执行原生 UI 冒烟测试
pwsh scripts/build-native.ps1 -SmokeTest
```

输出位于 `artifacts/win-x64`。WinUI XAML 编译器对较长的依赖路径有限制，如遇 WMC1006，请将 `NUGET_PACKAGES` 设置到较短的可写目录后重新还原。项目不会要求管理员权限，也不自动安装系统工作负载。

## 测试

- `PdfMailer.Tests`：原生 C# 业务回归，SMTP 全部模拟，不发送真实邮件。
- `PdfMailer.exe --smoke-test <输出目录>`：使用安全演示数据，实际启动 WinUI，检查原生筛选、勾选范围、长邮件滚动、工作表/模板/任务与多附件贯通、PDF 翻页缩放和附件切换，输出浅色、深色、设置页及预览 PNG 与 JSON 报告。
- CI：Windows/Linux 的 .NET 业务回归、Windows 原生构建/发布/UI 冒烟，以及保留的 Python CLI 回归。

## Python CLI 兼容入口

保留 `send_email.py` 及其回归供既有脚本使用；桌面版完全不调用它。

```powershell
python -m pip install -r requirements.txt
python send_email.py --config config.json -p "D:\PDF" --dry-run --export results.xlsx
python send_email.py --config config.json -p "D:\PDF" --state "D:\邮件任务\status.json"
python -m unittest -v test_send_email
```

CLI 使用 `--state` 指定日志，不自动读取原生配置的 `state_file` 字段。详细参数见 [配置指南](CONFIG_GUIDE.md)，操作步骤见 [使用说明](使用说明.md)。
