using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using MimeKit;
using Microsoft.Win32.SafeHandles;

namespace PdfMailer.Core;

public static partial class BatchPlanner
{
    [GeneratedRegex(@"^(?:签署后[：:]\s*)?(\d+)(?:_.+)?$", RegexOptions.CultureInvariant)]
    private static partial Regex FilenamePattern();
    [GeneratedRegex(@"^[^\s@<>,;]+@[^\s@<>,;]+\.[^\s@<>,;]+$", RegexOptions.CultureInvariant)]
    private static partial Regex EmailPattern();
    [GeneratedRegex(@"\{姓名\}|\{学号\}|姓名")]
    private static partial Regex PlaceholderPattern();

    public static string? ExtractStudentId(string filename)
    {
        var match = FilenamePattern().Match(Path.GetFileNameWithoutExtension(filename));
        return match.Success ? match.Groups[1].Value : null;
    }

    public static int ColumnNumber(string text)
    {
        text = text.Trim().ToUpperInvariant();
        if (int.TryParse(text, out var number))
        {
            if (number is >= 1 and <= 16384) return number;
        }
        else if (text.Length is > 0 and <= 3 && text.All(c => c is >= 'A' and <= 'Z'))
        {
            number = text.Aggregate(0, (value, c) => value * 26 + c - 'A' + 1);
            if (number <= 16384) return number;
        }
        throw new InvalidDataException("Excel 列应为 A 至 XFD 或 1 至 16384。");
    }

    public static bool ValidEmail(string value) => EmailPattern().IsMatch(value)
        && MailboxAddress.TryParse(value, out var address) && address.Address == value;

    public static void Validate(MailSettings settings, bool sending = false)
    {
        if (string.IsNullOrWhiteSpace(settings.SmtpServer)) throw new InvalidDataException("请输入 SMTP 服务器。");
        if (settings.SmtpPort is < 1 or > 65535) throw new InvalidDataException("SMTP 端口应为 1 至 65535。");
        if (settings.MaxRetries is < 1 or > 10) throw new InvalidDataException("最多尝试次数应为 1 至 10。");
        foreach (var interval in new[] { settings.SendInterval, settings.RetryInterval })
            if (!double.IsFinite(interval) || interval is < 0 or > 86400)
                throw new InvalidDataException("发送和重试间隔应为 0 至 86400 秒。");
        var columns = new[] { settings.ExcelStudentIdCol, settings.ExcelNameCol, settings.ExcelEmailCol };
        if (columns.Any(c => c is < 1 or > 16384) || columns.Distinct().Count() != 3)
            throw new InvalidDataException("学号、姓名和邮箱须使用不同的有效 Excel 列。");
        if (sending)
        {
            if (!ValidEmail(settings.SenderEmail)) throw new InvalidDataException("请输入有效的发件邮箱。");
            if (string.IsNullOrWhiteSpace(settings.SenderPassword)) throw new InvalidDataException("请输入邮箱授权码。");
            if (!settings.UseSsl && !settings.UseTls && settings.SmtpPort != 465)
                throw new InvalidDataException("发送必须启用 SSL 或 STARTTLS。");
        }
    }

    public static Dictionary<string, Student> LoadStudents(MailSettings settings)
    {
        using var workbook = new XLWorkbook(settings.StudentInfoExcel);
        var sheet = settings.ExcelWorksheet.Length > 0 ? workbook.Worksheets.FirstOrDefault(w => w.Name == settings.ExcelWorksheet)
            ?? throw new InvalidDataException("所选工作表已不存在，请重新导入。")
            : workbook.Worksheets.FirstOrDefault(w => w.TabActive) ?? workbook.Worksheet(1);
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        if (new[] { settings.ExcelStudentIdCol, settings.ExcelNameCol, settings.ExcelEmailCol }.Max() > lastColumn)
            throw new InvalidDataException("配置的列超出 Excel 的实际列数。");
        var students = new Dictionary<string, Student>(StringComparer.Ordinal);
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
        for (var row = 2; row <= lastRow; row++)
        {
            string Read(int column)
            {
                var cell = sheet.Cell(row, column);
                var value = cell.HasFormula ? cell.CachedValue : cell.Value;
                return value.ToString(CultureInfo.InvariantCulture).Trim();
            }
            var id = Read(settings.ExcelStudentIdCol);
            if (id.Length == 0) continue;
            if (!students.TryAdd(id, new(Read(settings.ExcelNameCol), Read(settings.ExcelEmailCol))))
                throw new InvalidDataException($"Excel 第 {row} 行学号重复：{id}。");
        }
        return students;
    }

    public static MailTemplate LoadTemplate(string path)
    {
        var content = File.ReadAllText(path, Encoding.UTF8).TrimStart('\uFEFF')
            .Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Trim();
        var split = content.IndexOf('\n');
        if (split < 0) throw new InvalidDataException("模板须包含主题和正文，第一行为主题。");
        var subject = content[..split].Trim();
        if (subject.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase)) subject = subject[8..].Trim();
        var body = content[(split + 1)..].Trim();
        if (subject.Length == 0 || body.Length == 0) throw new InvalidDataException("邮件主题和正文不能为空。");
        return new(subject, body);
    }

    public static string Personalize(string text, string name, string id) =>
        PlaceholderPattern().Replace(text, match => match.Value == "{学号}" ? id : name);

    public static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexStringLower(SHA256.HashData(stream));
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(SafeFileHandle file, StringBuilder path, uint length, uint flags);

    public static string CanonicalPath(string path)
    {
        if (!OperatingSystem.IsWindows()) return Path.GetFullPath(path);
        using var handle = File.OpenHandle(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var buffer = new StringBuilder(32768);
        var length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
        if (length == 0 || length >= buffer.Capacity) throw new IOException("无法解析附件的完整路径。");
        var canonical = buffer.ToString();
        return canonical.StartsWith(@"\\?\UNC\", StringComparison.Ordinal) ? @"\\" + canonical[8..]
            : canonical.StartsWith(@"\\?\", StringComparison.Ordinal) ? canonical[4..] : canonical;
    }

    // 保持与 Python v2 json.dumps(..., ensure_ascii=False) 的身份编码相同。
    public static string Identity(params string[] values)
    {
        static string Quote(string value)
        {
            var output = new StringBuilder("\"");
            foreach (var c in value)
                output.Append(c switch
                {
                    '"' => "\\\"", '\\' => "\\\\", '\b' => "\\b", '\f' => "\\f", '\n' => "\\n",
                    '\r' => "\\r", '\t' => "\\t", < ' ' => "\\u" + ((int)c).ToString("x4"), _ => c.ToString()
                });
            return output.Append('"').ToString();
        }
        return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes("[" + string.Join(", ", values.Select(Quote)) + "]")));
    }

    public static IReadOnlyList<MailJob> Prepare(MailSettings settings, string journalPath, CancellationToken cancellation = default)
    {
        Validate(settings);
        if (!Directory.Exists(settings.AttachmentFolder)) throw new DirectoryNotFoundException("请选择存在的 PDF 附件目录。");
        var students = LoadStudents(settings);
        var template = LoadTemplate(settings.EmailTemplate);
        var journal = Storage.LoadJournal(journalPath);
        var jobs = new List<MailJob>();
        Span<byte> header = stackalloc byte[5];
        foreach (var path in Directory.EnumerateFiles(Path.GetFullPath(settings.AttachmentFolder)).Order(StringComparer.Ordinal))
        {
            cancellation.ThrowIfCancellationRequested();
            if (!Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase)) continue;
            var job = new MailJob { Filename = Path.GetFileName(path), Filepath = path, StudentId = ExtractStudentId(path) ?? "" };
            try
            {
                job.Filepath = CanonicalPath(path);
                if (job.StudentId.Length == 0) throw new InvalidDataException("文件名无法提取学号。");
                if (!students.TryGetValue(job.StudentId, out var student)) throw new InvalidDataException("Excel 中找不到此学号。");
                job.Name = student.Name;
                job.Email = student.Email;
                if (job.Name.Length == 0) throw new InvalidDataException("姓名为空。");
                if (!ValidEmail(job.Email)) throw new InvalidDataException("收件邮箱格式无效。");
                using (var stream = File.OpenRead(path))
                {
                    if (stream.Read(header) != 5 || !header.SequenceEqual("%PDF-"u8))
                        throw new InvalidDataException("文件缺少 PDF 标识，请检查附件。");
                }
                job.Digest = HashFile(path);
                job.Subject = Personalize(template.Subject, job.Name, job.StudentId);
                job.Body = Personalize(template.Body, job.Name, job.StudentId);
                if (job.Subject.ContainsAny('\r', '\n')) throw new InvalidDataException("主题含有换行，请检查姓名和模板。");
                job.Key = Identity(job.Filepath, job.Digest, settings.SenderEmail.Trim().ToLowerInvariant(), job.Email, job.Subject, job.Body);
                if (journal.Records.TryGetValue(job.Key, out var previous))
                {
                    job.Status = previous.Status == JobStates.Sending ? JobStates.Uncertain : previous.Status;
                    job.Error = previous.Status == JobStates.Sending ? "上次提交中断，请核实已发送邮件。" : previous.Error;
                }
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException)
            { job.Status = JobStates.Blocked; job.Error = exception.Message; }
            jobs.Add(job);
        }
        return settings.GroupAttachments ? MergeAttachments(jobs, settings, journal) : jobs;
    }

    private static IReadOnlyList<MailJob> MergeAttachments(List<MailJob> jobs, MailSettings settings, SendJournal journal)
    {
        var result = new List<MailJob>(jobs.Where(j => j.Status == JobStates.Blocked));
        foreach (var group in jobs.Where(j => j.Status != JobStates.Blocked).GroupBy(j => (j.StudentId, j.Email)))
        {
            var items = group.ToArray();
            if (items.Length == 1) { result.Add(items[0]); continue; }
            var first = items[0];
            var files = items.SelectMany(j => j.Files.Select(f => f with { IdentityKey = j.Key })).ToArray();
            var key = Identity(["bundle-v1", settings.SenderEmail.Trim().ToLowerInvariant(), first.Email, first.Subject, first.Body,
                .. files.SelectMany(f => new[] { f.Filepath, f.Digest })]);
            var merged = new MailJob { Filename = string.Join("、", files.Select(f => f.Filename)), Filepath = first.Filepath,
                StudentId = first.StudentId, Name = first.Name, Email = first.Email, Subject = first.Subject, Body = first.Body,
                Key = key, Attachments = files };
            if (journal.Records.TryGetValue(key, out var previous))
            {
                merged.Status = previous.Status == JobStates.Sending ? JobStates.Uncertain : previous.Status;
                merged.Error = previous.Error;
            }
            else if (items.All(j => j.Status == JobStates.Success)) merged.Status = JobStates.Success;
            else if (items.Any(j => j.Status is JobStates.Success or JobStates.Uncertain or JobStates.Sending))
            {
                merged.Status = JobStates.Blocked;
                merged.Error = "这组附件已有部分发送或待核实记录。请关闭合并模式，核对并处理剩余附件，避免重复发送。";
            }
            result.Add(merged);
        }
        return result.OrderBy(j => j.Filename, StringComparer.Ordinal).ToArray();
    }
}
