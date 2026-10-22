using ClosedXML.Excel;
using MailKit.Security;

namespace PdfMailer.Core;

public sealed record WorksheetPreview(string Name, IReadOnlyList<string> Headers, IReadOnlyList<string[]> Rows,
    int IdColumn, int NameColumn, int EmailColumn);

public static class ExcelImport
{
    public static IReadOnlyList<WorksheetPreview> Inspect(string path)
    {
        using var book = new XLWorkbook(path);
        return book.Worksheets.Select(sheet =>
        {
            var width = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            var headers = Enumerable.Range(1, width).Select(c => sheet.Cell(1, c).GetString().Trim()).ToArray();
            int Find(params string[] names) => Array.FindIndex(headers, h => names.Contains(h, StringComparer.OrdinalIgnoreCase)) + 1;
            var rows = Enumerable.Range(2, Math.Max(0, Math.Min(5, (sheet.LastRowUsed()?.RowNumber() ?? 1) - 1)))
                .Select(r => Enumerable.Range(1, width).Select(c => sheet.Cell(r, c).HasFormula
                    ? sheet.Cell(r, c).CachedValue.ToString() : sheet.Cell(r, c).Value.ToString()).ToArray()).ToArray();
            return new WorksheetPreview(sheet.Name, headers, rows, Find("学号", "学生编号", "studentid", "student_id", "id"),
                Find("姓名", "学生姓名", "name"), Find("邮箱", "电子邮箱", "电子邮件", "email", "e-mail"));
        }).ToArray();
    }
}

public enum SendScope { All, Included, Visible }
public static class SendSelection
{
    public static IReadOnlyList<MailJob> Select(IReadOnlyList<MailJob> jobs, IReadOnlyList<MailJob> visible, SendScope scope, bool retryOnly = false)
    {
        var candidates = scope == SendScope.Visible ? visible : jobs;
        return candidates.Where(j => (scope != SendScope.Included || j.IsIncluded)
            && (retryOnly ? j.Status == JobStates.Failed : JobStates.CanSend(j.Status))).ToArray();
    }
}

public sealed record SavedTask(string Id, string Name, DateTimeOffset UpdatedAt, MailSettings Settings,
    int Total, int Success, int Attention);
public sealed record SavedTemplate(string Id, string Name, string Subject, string Body);
public sealed record SavedAccount(string Email, string Server, int Port, bool UseSsl, bool RememberPassword);

public sealed class WorkspaceStore(string directory)
{
    private string PathFor(string name) => Path.Combine(directory, name + ".json");
    private List<T> Read<T>(string name) => File.Exists(PathFor(name))
        ? System.Text.Json.JsonSerializer.Deserialize<List<T>>(File.ReadAllText(PathFor(name)), Storage.Json)
            ?? throw new InvalidDataException("保存的资料为空：" + name) : [];
    public IReadOnlyList<SavedTask> Tasks => Read<SavedTask>("tasks").OrderByDescending(t => t.UpdatedAt).ToArray();
    public IReadOnlyList<SavedTemplate> Templates => Read<SavedTemplate>("templates");
    public IReadOnlyList<SavedAccount> Accounts => Read<SavedAccount>("accounts");
    public void SaveTask(SavedTask task)
    {
        var values = Read<SavedTask>("tasks"); values.RemoveAll(t => t.Id == task.Id);
        values.Add(task with { Settings = task.Settings with { SenderPassword = "" } }); Storage.WriteAtomic(PathFor("tasks"), values);
    }
    public void SaveTemplate(SavedTemplate template)
    {
        if (string.IsNullOrWhiteSpace(template.Name) || string.IsNullOrWhiteSpace(template.Subject) || string.IsNullOrWhiteSpace(template.Body)
            || template.Subject.ContainsAny('\r', '\n')) throw new InvalidDataException("请输入模板名称、单行主题和正文。");
        var values = Read<SavedTemplate>("templates"); values.RemoveAll(t => t.Id == template.Id);
        values.Add(template); Storage.WriteAtomic(PathFor("templates"), values);
    }
    public string Materialize(SavedTemplate template)
    {
        var folder = Path.Combine(directory, "templates"); Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, template.Subject + "\n\n" + template.Body, new System.Text.UTF8Encoding(false));
        return path;
    }
    public void SaveAccount(SavedAccount account)
    {
        var values = Read<SavedAccount>("accounts"); values.RemoveAll(a => a.Email.Equals(account.Email, StringComparison.OrdinalIgnoreCase));
        values.Add(account); Storage.WriteAtomic(PathFor("accounts"), values);
    }
    public void DeleteAccount(string email)
    {
        var values = Read<SavedAccount>("accounts"); values.RemoveAll(a => a.Email.Equals(email, StringComparison.OrdinalIgnoreCase));
        Storage.WriteAtomic(PathFor("accounts"), values);
    }
}

public sealed class AccountDiagnostics(Func<ISmtpSession>? factory = null)
{
    public async Task CheckAsync(MailSettings settings, bool authenticate)
    {
        BatchPlanner.Validate(settings, authenticate);
        var session = factory?.Invoke() ?? new MailKitSession();
        try
        {
            await session.ConnectAsync(settings.SmtpServer, settings.SmtpPort,
                settings.UseSsl || settings.SmtpPort == 465 ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTls).ConfigureAwait(false);
            if (authenticate) await session.AuthenticateAsync(settings.SenderEmail, settings.SenderPassword).ConfigureAwait(false);
        }
        finally
        {
            try { await session.DisconnectAsync().ConfigureAwait(false); } catch { }
            try { session.Dispose(); } catch { }
        }
    }

    public static MailJob CreateTest(MailJob original, string ownAddress)
    {
        if (!BatchPlanner.ValidEmail(ownAddress)) throw new InvalidDataException("试发地址无效。");
        return new MailJob { Filename = original.Filename, Filepath = original.Filepath, Digest = original.Digest,
            Attachments = original.Files.Select(f => f with { IdentityKey = "" }).ToArray(),
            StudentId = original.StudentId, Name = original.Name, Email = ownAddress,
            Subject = "[试发] " + original.Subject, Body = original.Body, Key = Guid.NewGuid().ToString("N") };
    }
}
