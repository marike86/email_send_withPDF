using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace PdfMailer.Core;

public static class Storage
{
    public static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public static void WriteAtomic<T>(string path, T value)
    {
        path = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                JsonSerializer.Serialize(stream, value, Json);
                stream.Flush(true);
            }
            File.Move(temporary, path, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    public static MailSettings LoadSettings(string path)
    {
        var settings = new MailSettings();
        if (File.Exists(path))
        {
            using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                throw new InvalidDataException("配置文件必须是 JSON 对象");
            // 兼容 Python 版以字母表达的列，统一后再反序列化。
            var values = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(document.RootElement.GetRawText())!;
            foreach (var key in new[] { "excel_student_id_col", "excel_name_col", "excel_email_col" })
                if (values.TryGetValue(key, out var element) && element.ValueKind == JsonValueKind.String)
                    values[key] = JsonSerializer.SerializeToElement(BatchPlanner.ColumnNumber(element.GetString()!));
            settings = JsonSerializer.Deserialize<MailSettings>(JsonSerializer.Serialize(values), Json)
                ?? throw new InvalidDataException("配置内容为空");
            // 最早的 Python 配置通常只有 use_tls，不能套用新建账户的 SSL 默认值。
            if (!document.RootElement.TryGetProperty("use_ssl", out _)) settings.UseSsl = settings.SmtpPort == 465;
            if (!document.RootElement.TryGetProperty("use_tls", out _)) settings.UseTls = !settings.UseSsl;
            if (document.RootElement.TryGetProperty("sender_password", out var password))
                settings.SenderPassword = password.GetString() ?? ""; // 仅兼容读取，不保存。
        }
        var type = typeof(MailSettings);
        foreach (var property in type.GetProperties())
        {
            var name = JsonNamingPolicy.SnakeCaseLower.ConvertName(property.Name).ToUpperInvariant();
            var raw = Environment.GetEnvironmentVariable(name);
            if (raw is null) continue;
            object value = property.PropertyType == typeof(string) ? raw
                : property.PropertyType == typeof(bool) ? raw.ToLowerInvariant() switch
                {
                    "true" or "1" => true, "false" or "0" => false,
                    _ => throw new InvalidDataException($"{name} 应为 true/false")
                }
                : property.Name.StartsWith("Excel", StringComparison.Ordinal) ? BatchPlanner.ColumnNumber(raw)
                : property.PropertyType == typeof(int) ? int.Parse(raw, CultureInfo.InvariantCulture)
                : double.Parse(raw, CultureInfo.InvariantCulture);
            property.SetValue(settings, value);
        }
        var directory = Path.GetDirectoryName(Path.GetFullPath(path))!;
        settings.AttachmentFolder = Resolve(settings.AttachmentFolder, directory);
        settings.StudentInfoExcel = Resolve(settings.StudentInfoExcel, directory);
        settings.EmailTemplate = Resolve(settings.EmailTemplate, directory);
        settings.StateFile = Resolve(settings.StateFile, directory);
        return settings;
    }

    public static string Resolve(string path, string directory) => string.IsNullOrWhiteSpace(path) ? ""
        : Path.GetFullPath(path.Trim().Trim('"', '\''), directory);

    public static SendJournal LoadJournal(string path)
    {
        if (!File.Exists(path)) return new();
        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object || !root.TryGetProperty("version", out var version)
            || !version.TryGetInt32(out var number) || number != 2
            || !root.TryGetProperty("records", out var records) || records.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("日志为旧版或格式损坏，请备份核对后选择新的发送日志；不会自动清空历史。");
        var journal = JsonSerializer.Deserialize<SendJournal>(root.GetRawText(), Json)!;
        if (journal.Records is null || journal.Records.Values.Any(r => r is null || r.Status is not
            (JobStates.Success or JobStates.Failed or JobStates.Sending or JobStates.Uncertain)))
            throw new InvalidDataException("发送日志包含无效记录，请备份核对。");
        return journal;
    }

    public static FileStream AcquireJournal(string path)
    {
        var lockPath = Path.GetFullPath(path) + ".lock";
        Directory.CreateDirectory(Path.GetDirectoryName(lockPath)!);
        try
        {
            // Windows 上 FileShare.None 还阻止 Python 版同时持有此日志。
            return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        }
        catch (IOException exception)
        { throw new IOException("另一个任务正在使用这份发送日志，请等待其结束。", exception); }
    }

    public static void ResolveUncertain(string path, string key, bool delivered, IEnumerable<string>? relatedKeys = null)
    {
        using var journalLock = AcquireJournal(path);
        var journal = LoadJournal(path);
        if (!journal.Records.TryGetValue(key, out var record) || record.Status is not (JobStates.Sending or JobStates.Uncertain))
            throw new InvalidOperationException("记录已经变化，请重新检查文件。");
        record.Status = delivered ? JobStates.Success : JobStates.Failed;
        record.Error = "用户已核实发送结果";
        record.Time = DateTimeOffset.Now.ToString("O");
        foreach (var related in relatedKeys ?? [])
            if (journal.Records.TryGetValue(related, out var child) && child.Status is JobStates.Sending or JobStates.Uncertain)
            {
                child.Status = record.Status; child.Error = record.Error; child.Time = record.Time;
            }
        WriteAtomic(path, journal);
    }
}
