using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using PdfMailer.Core;
using Xunit;

namespace PdfMailer.Tests;

public sealed class BatchTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "pdfmailer-tests-" + Guid.NewGuid().ToString("N"));
    private readonly MailSettings settings;
    private readonly string state;
    private string Pdf => Path.Combine(settings.AttachmentFolder, "001_证书.PDF");
    public BatchTests()
    {
        Directory.CreateDirectory(root);
        settings = new()
        {
            AttachmentFolder = Path.Combine(root, "pdf"), StudentInfoExcel = Path.Combine(root, "students.xlsx"),
            EmailTemplate = Path.Combine(root, "template.txt"), SenderEmail = "sender@example.com", SenderPassword = "secret",
            SendInterval = 0, RetryInterval = 0
        };
        Directory.CreateDirectory(settings.AttachmentFolder);
        File.WriteAllBytes(Pdf, "%PDF-1.4\nfixture"u8.ToArray());
        File.WriteAllText(settings.EmailTemplate, "Subject: {姓名}的资料\r\n\r\n{姓名}同学，学号 {学号}", new UTF8Encoding(true));
        WriteStudents(("001", "张三", "student@example.com"));
        state = Path.Combine(root, "status.json");
    }
    public void Dispose() => Directory.Delete(root, true);
    private void WriteStudents(params (string Id, string Name, string Email)[] rows)
    {
        using var book = new XLWorkbook();
        var sheet = book.AddWorksheet("学生");
        sheet.Cell(1, 1).Value = "学号"; sheet.Cell(1, 2).Value = "姓名"; sheet.Cell(1, 3).Value = "邮箱";
        for (var row = 0; row < rows.Length; row++)
        { sheet.Cell(row + 2, 1).Value = rows[row].Id; sheet.Cell(row + 2, 2).Value = rows[row].Name; sheet.Cell(row + 2, 3).Value = rows[row].Email; }
        book.SaveAs(settings.StudentInfoExcel);
    }
    private IReadOnlyList<MailJob> Prepare() => BatchPlanner.Prepare(settings, state);
    private static void Apply(IReadOnlyList<MailJob> jobs, IReadOnlyList<JobUpdate> updates)
    { foreach (var update in updates) { var job = jobs.First(j => j.Key == update.Key); job.Status = update.Status; job.Error = update.Error; } }

    [Fact] public void PreflightPersonalizesOffline()
    {
        var job = Assert.Single(Prepare());
        Assert.Equal(JobStates.Ready, job.Status); Assert.Equal("张三的资料", job.Subject);
        Assert.Contains("001", job.Body); Assert.DoesNotContain('\r', job.Body); Assert.False(File.Exists(state));
    }
    [Theory]
    [InlineData("2023001_资料.pdf", "2023001")]
    [InlineData("签署后：2023001_001.pdf", "2023001")]
    [InlineData("001.pdf", "001")]
    [InlineData("别人的123_文件.pdf", null)]
    public void FilenameMatchesWholeId(string name, string? expected) => Assert.Equal(expected, BatchPlanner.ExtractStudentId(name));
    [Theory]
    [InlineData("A", 1)] [InlineData("AA", 27)] [InlineData("XFD", 16384)] [InlineData("27", 27)]
    public void MultiLetterColumns(string text, int number) => Assert.Equal(number, BatchPlanner.ColumnNumber(text));
    [Theory] [InlineData("0")] [InlineData("XFE")] [InlineData("-1")]
    public void InvalidColumnsRejected(string text) => Assert.Throws<InvalidDataException>(() => BatchPlanner.ColumnNumber(text));
    [Theory]
    [InlineData("a@example.com", true)] [InlineData("a@example.com,b@example.com", false)]
    [InlineData("a@example.com\r\nBcc: b@example.com", false)] [InlineData("", false)]
    public void SingleMailboxOnly(string address, bool valid) => Assert.Equal(valid, BatchPlanner.ValidEmail(address));
    [Fact] public void DuplicateStudentsRejected()
    { WriteStudents(("001", "A", "a@example.com"), ("001", "B", "b@example.com")); Assert.Throws<InvalidDataException>(() => Prepare()); }
    [Fact] public void UnmatchedAndInvalidPdfsAreVisible()
    {
        File.WriteAllText(Path.Combine(settings.AttachmentFolder, "unmatched.pdf"), "%PDF-test");
        File.WriteAllText(Path.Combine(settings.AttachmentFolder, "001_invalid.pdf"), "not PDF");
        File.WriteAllText(Path.Combine(settings.AttachmentFolder, "001_private.txt"), "secret");
        var jobs = Prepare(); Assert.Equal(3, jobs.Count); Assert.Equal(2, jobs.Count(j => j.Status == JobStates.Blocked));
    }
    [Fact] public void InvalidRecipientBlocked()
    { WriteStudents(("001", "A", "not-address")); Assert.Equal(JobStates.Blocked, Assert.Single(Prepare()).Status); }
    [Fact] public void LockedAttachmentIsVisibleAsBlocked()
    {
        using var held = new FileStream(Pdf, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.Equal(JobStates.Blocked, Assert.Single(Prepare()).Status);
    }
    [Fact] public void TemplateReplacementDoesNotInterpretInsertedName()
    { Assert.Equal("姓名{学号} / 001", BatchPlanner.Personalize("{姓名} / {学号}", "姓名{学号}", "001")); }

    [Fact] public async Task SuccessSkippedOnResume()
    {
        var transport = new FakeTransport(); var sender = new BatchSender(transport);
        await sender.RunAsync(settings, Prepare(), state); await sender.RunAsync(settings, Prepare(), state);
        Assert.Equal(1, transport.Calls); Assert.Equal(JobStates.Success, Assert.Single(Prepare()).Status);
    }
    [Fact] public async Task StalePreflightRechecksJournal()
    {
        var jobs = Prepare(); var stale = Prepare(); var transport = new FakeTransport(); var sender = new BatchSender(transport);
        await sender.RunAsync(settings, jobs, state); await sender.RunAsync(settings, stale, state);
        Assert.Equal(1, transport.Calls);
    }
    [Fact] public async Task ChangedAttachmentNeverSends()
    {
        var jobs = Prepare(); File.AppendAllText(Pdf, "changed"); var transport = new FakeTransport();
        var results = await new BatchSender(transport).RunAsync(settings, jobs, state);
        Assert.Equal(0, transport.Calls); Assert.Equal(JobStates.Failed, Assert.Single(results).Status);
        Assert.Equal(JobStates.Ready, Assert.Single(Prepare()).Status);
    }
    [Fact] public async Task FailedThenSuccessReplacesRecord()
    {
        var transport = new FakeTransport { NextError = new AuthenticationException("denied") };
        var sender = new BatchSender(transport);
        await sender.RunAsync(settings, Prepare(), state); Assert.Equal(1, transport.Calls);
        Assert.Equal(JobStates.Failed, Assert.Single(Prepare()).Status);
        await sender.RunAsync(settings, Prepare(), state, retryOnly: true);
        Assert.Equal(JobStates.Success, Assert.Single(Storage.LoadJournal(state).Records).Value.Status);
    }
    [Fact] public async Task UncertainNeverAutomaticallyRetried()
    {
        var transport = new FakeTransport { NextError = new DeliveryUncertainException("verify") };
        var sender = new BatchSender(transport);
        await sender.RunAsync(settings, Prepare(), state); await sender.RunAsync(settings, Prepare(), state);
        Assert.Equal(1, transport.Calls); Assert.Equal(JobStates.Uncertain, Assert.Single(Prepare()).Status);
    }
    [Fact] public void InterruptedSendingRequiresVerification()
    {
        var job = Assert.Single(Prepare());
        Storage.WriteAtomic(state, new SendJournal { Records = new() { [job.Key] = JobRecord.From(job) with { Status = JobStates.Sending } } });
        Assert.Equal(JobStates.Uncertain, Assert.Single(Prepare()).Status);
        Storage.ResolveUncertain(state, job.Key, true); Assert.Equal(JobStates.Success, Assert.Single(Prepare()).Status);
    }
    [Fact] public async Task StopBeforeSend()
    { using var stop = new CancellationTokenSource(); stop.Cancel(); var transport = new FakeTransport(); await new BatchSender(transport).RunAsync(settings, Prepare(), state, stop: stop.Token); Assert.Equal(0, transport.Calls); }
    [Fact] public async Task StopAfterCurrentMessage()
    {
        File.WriteAllText(Path.Combine(settings.AttachmentFolder, "001_second.pdf"), "%PDF-test");
        using var stop = new CancellationTokenSource(); var transport = new FakeTransport { OnSend = () => stop.Cancel() };
        var result = await new BatchSender(transport).RunAsync(settings, Prepare(), state, stop: stop.Token);
        Assert.Single(result); Assert.Equal(JobStates.Success, result[0].Status); Assert.Equal(1, transport.Calls);
    }
    [Fact] public async Task TransientConnectionErrorRetries()
    { var transport = new FakeTransport { NextError = new IOException("offline") }; await new BatchSender(transport).RunAsync(settings, Prepare(), state); Assert.Equal(2, transport.Calls); }
    [Theory] [InlineData("{bad json")] [InlineData("{\"processed_files\":[]}")] [InlineData("{\"version\":2,\"records\":{\"bad\":{\"status\":\"oops\"}}}")]
    public void CorruptAndLegacyJournalNeverReset(string text)
    { File.WriteAllText(state, text); Assert.ThrowsAny<Exception>(() => Prepare()); Assert.Equal(text, File.ReadAllText(state)); }
    [Fact] public async Task LockedJournalPreventsNetwork()
    {
        var jobs = Prepare(); using var held = Storage.AcquireJournal(state); var transport = new FakeTransport();
        await Assert.ThrowsAsync<IOException>(() => new BatchSender(transport).RunAsync(settings, jobs, state)); Assert.Equal(0, transport.Calls);
    }
    [Fact] public async Task JournalWriteFailurePreventsNetwork()
    {
        var jobs = Prepare(); Directory.CreateDirectory(state); var transport = new FakeTransport();
        var error = await Record.ExceptionAsync(() => new BatchSender(transport).RunAsync(settings, jobs, state));
        Assert.True(error is IOException or UnauthorizedAccessException); Assert.Equal(0, transport.Calls);
    }
    [Fact] public void PasswordNotSavedAndPathsResolve()
    {
        var config = Path.Combine(root, "config.json"); Storage.WriteAtomic(config, settings with { EmailTemplate = "template.txt" });
        var text = File.ReadAllText(config); Assert.DoesNotContain("secret", text); Assert.DoesNotContain("sender_password", text);
        Assert.Equal(Path.GetFullPath(settings.EmailTemplate), Storage.LoadSettings(config).EmailTemplate);
    }
    [Fact] public void OldPlaintextPasswordReadButNotRewritten()
    {
        var config = Path.Combine(root, "config.json"); File.WriteAllText(config, "{\"sender_password\":\"legacy\",\"excel_email_col\":\"AA\"}");
        var value = Storage.LoadSettings(config); Assert.Equal("legacy", value.SenderPassword); Assert.Equal(27, value.ExcelEmailCol);
        Storage.WriteAtomic(config, value); Assert.DoesNotContain("legacy", File.ReadAllText(config));
    }
    [Fact] public void LegacyStartTlsConfigDoesNotInheritSslDefault()
    {
        var path = Path.Combine(root, "config.json");
        File.WriteAllText(path, "{\"smtp_port\":587,\"use_tls\":true}");
        var loaded = Storage.LoadSettings(path);
        Assert.True(loaded.UseTls); Assert.False(loaded.UseSsl); Assert.Equal(587, loaded.SmtpPort);
    }
    [Fact] public void PythonV2IdentityIsCompatible()
    {
        // 参考值由 Python json.dumps(..., ensure_ascii=False) + hashlib.sha256 生成。
        var identity = BatchPlanner.Identity(@"D:\资料\001_证书.pdf", "abc123", "sender@example.com", "student@example.com", "张三的资料", "你好\n正文");
        Assert.Equal("ef47485cff1533680c035074000909c03b9faafd3764aa6fce85edb5b561fbf7", identity);
    }
    [Fact] public void WindowsPathCasingDoesNotCreateNewJob()
    {
        if (!OperatingSystem.IsWindows()) return;
        var original = Assert.Single(Prepare());
        var altered = Assert.Single(BatchPlanner.Prepare(settings with { AttachmentFolder = settings.AttachmentFolder.ToUpperInvariant() }, state));
        Assert.Equal(original.Key, altered.Key);
    }
    [Fact] public void ExportXlsxPreservesIdsAndEscapesFormula()
    {
        var jobs = Prepare(); jobs[0].Name = "=1+1"; var path = Path.Combine(root, "result.xlsx"); ResultExporter.Export(jobs, path);
        using var book = new XLWorkbook(path); Assert.Equal("001", book.Worksheet(1).Cell(2, 2).GetString());
        var name = book.Worksheet(1).Cell(2, 3);
        Assert.False(name.HasFormula); Assert.Equal("=1+1", name.GetString()); Assert.True(name.Style.IncludeQuotePrefix);
    }
    [Fact] public void ExportCsvHasBomAndQuotes()
    {
        var jobs = Prepare(); jobs[0].Name = "=1+1"; var path = Path.Combine(root, "result.csv"); ResultExporter.Export(jobs, path);
        Assert.True(File.ReadAllBytes(path).AsSpan().StartsWith(new byte[] { 0xef, 0xbb, 0xbf })); Assert.Contains("'=1+1", File.ReadAllText(path));
    }
    [Theory] [InlineData(465, true, false, SecureSocketOptions.SslOnConnect)] [InlineData(587, false, true, SecureSocketOptions.StartTls)]
    public async Task TransportUsesExplicitSecurity(int port, bool ssl, bool tls, SecureSocketOptions expected)
    {
        var session = new FakeSession();
        await new SmtpTransport(() => session).SendAsync(settings with { SmtpPort = port, UseSsl = ssl, UseTls = tls }, Assert.Single(Prepare()));
        Assert.Equal(expected, session.Security); Assert.Equal("application/pdf", session.AttachmentType); Assert.Equal("001_证书.PDF", session.AttachmentName);
    }
    [Fact] public async Task DisconnectDuringSubmissionIsUncertain()
    {
        var session = new FakeSession { SendError = new IOException("lost") };
        await Assert.ThrowsAsync<DeliveryUncertainException>(() => new SmtpTransport(() => session).SendAsync(settings, Assert.Single(Prepare())));
    }
    [Fact] public async Task CleanupFailureDoesNotChangeSuccess()
    {
        var session = new FakeSession { CleanupError = true };
        await new SmtpTransport(() => session).SendAsync(settings, Assert.Single(Prepare())); Assert.True(session.Sent);
    }
    [Fact] public async Task MissingAttachmentDoesNotConnect()
    {
        var job = Assert.Single(Prepare()); File.Delete(Pdf); var session = new FakeSession();
        await Assert.ThrowsAsync<FileNotFoundException>(() => new SmtpTransport(() => session).SendAsync(settings, job)); Assert.False(session.Connected);
    }
    [Fact] public async Task TransportRechecksActualBytes()
    {
        var job = Assert.Single(Prepare()); File.AppendAllText(Pdf, "changed"); var session = new FakeSession();
        await Assert.ThrowsAsync<InvalidDataException>(() => new SmtpTransport(() => session).SendAsync(settings, job)); Assert.False(session.Connected);
    }
    [Fact] public void InsecureAndInvalidSettingsRejected()
    {
        Assert.Throws<InvalidDataException>(() => BatchPlanner.Validate(settings with { SmtpPort = 25, UseSsl = false, UseTls = false }, true));
        Assert.Throws<InvalidDataException>(() => BatchPlanner.Validate(settings with { SendInterval = double.NaN }));
        Assert.Throws<InvalidDataException>(() => BatchPlanner.Validate(settings with { MaxRetries = 0 }));
    }

    private sealed class FakeTransport : IMailTransport
    {
        public int Calls; public Exception? NextError; public Action? OnSend;
        public Task SendAsync(MailSettings settings, MailJob job)
        { Calls++; OnSend?.Invoke(); if (NextError is { } error) { NextError = null; return Task.FromException(error); } return Task.CompletedTask; }
    }
    private sealed class FakeSession : ISmtpSession
    {
        public SecureSocketOptions Security; public bool Connected; public bool Sent; public Exception? SendError; public bool CleanupError;
        public string? AttachmentName; public string? AttachmentType;
        public Task ConnectAsync(string host, int port, SecureSocketOptions security) { Security = security; Connected = true; return Task.CompletedTask; }
        public Task AuthenticateAsync(string user, string password) => Task.CompletedTask;
        public Task SendAsync(MimeMessage message)
        {
            if (SendError is not null) return Task.FromException(SendError);
            var attachment = Assert.Single(message.Attachments); AttachmentName = ((MimePart)attachment).FileName;
            AttachmentType = attachment.ContentType.MimeType; Sent = true; return Task.CompletedTask;
        }
        public Task DisconnectAsync() => CleanupError ? Task.FromException(new IOException("cleanup")) : Task.CompletedTask;
        public void Dispose() { if (CleanupError) throw new IOException("dispose"); }
    }
}
