using System.Security.Cryptography;
using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace PdfMailer.Core;

public interface IMailTransport
{
    Task SendAsync(MailSettings settings, MailJob job);
}

public interface ISmtpSession : IDisposable
{
    Task ConnectAsync(string host, int port, SecureSocketOptions security);
    Task AuthenticateAsync(string user, string password);
    Task SendAsync(MimeMessage message);
    Task DisconnectAsync();
}

public sealed class MailKitSession : ISmtpSession
{
    private readonly SmtpClient client = new() { Timeout = 30000 };
    public Task ConnectAsync(string host, int port, SecureSocketOptions security) => client.ConnectAsync(host, port, security);
    public Task AuthenticateAsync(string user, string password) => client.AuthenticateAsync(user, password);
    public async Task SendAsync(MimeMessage message) => await client.SendAsync(message).ConfigureAwait(false);
    public Task DisconnectAsync() => client.IsConnected ? client.DisconnectAsync(true) : Task.CompletedTask;
    public void Dispose() => client.Dispose();
}

public sealed class SmtpTransport(Func<ISmtpSession>? factory = null) : IMailTransport
{
    public async Task SendAsync(MailSettings settings, MailJob job)
    {
        BatchPlanner.Validate(settings, true);
        if (!BatchPlanner.ValidEmail(job.Email)) throw new InvalidDataException("收件邮箱格式无效。");
        var bytes = await File.ReadAllBytesAsync(job.Filepath).ConfigureAwait(false);
        if (!bytes.AsSpan().StartsWith("%PDF-"u8)) throw new InvalidDataException("文件缺少 PDF 标识。");
        if (Convert.ToHexStringLower(SHA256.HashData(bytes)) != job.Digest)
            throw new InvalidDataException("附件在预检后发生变化，请重新检查文件。");
        using var message = new MimeMessage();
        message.From.Add(MailboxAddress.Parse(settings.SenderEmail));
        message.To.Add(MailboxAddress.Parse(job.Email));
        message.Subject = job.Subject;
        message.MessageId = MimeKit.Utils.MimeUtils.GenerateMessageId();
        var builder = new BodyBuilder { TextBody = job.Body };
        builder.Attachments.Add(job.Filename, bytes, new ContentType("application", "pdf"));
        message.Body = builder.ToMessageBody();
        var session = factory?.Invoke() ?? new MailKitSession();
        try
        {
            await session.ConnectAsync(settings.SmtpServer, settings.SmtpPort,
                settings.UseSsl || settings.SmtpPort == 465 ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTls).ConfigureAwait(false);
            await session.AuthenticateAsync(settings.SenderEmail, settings.SenderPassword).ConfigureAwait(false);
            try { await session.SendAsync(message).ConfigureAwait(false); }
            catch (Exception exception) when (exception is IOException or SmtpProtocolException or TimeoutException or OperationCanceledException)
            { throw new DeliveryUncertainException("提交邮件时连接中断，可能已经送达，请先核实。", exception); }
        }
        finally
        {
            // 连接清理失败不得将已成功提交的邮件重新计作失败。
            try { await session.DisconnectAsync().ConfigureAwait(false); } catch { }
            try { session.Dispose(); } catch { }
        }
    }
}

public sealed class BatchSender(IMailTransport? transport = null)
{
    private readonly IMailTransport transport = transport ?? new SmtpTransport();

    private static bool IsTransient(Exception exception) => exception switch
    {
        AuthenticationException or System.Security.Authentication.AuthenticationException => false,
        FileNotFoundException or DirectoryNotFoundException or UnauthorizedAccessException or InvalidDataException => false,
        SmtpCommandException response => (int)response.StatusCode is >= 400 and < 500,
        IOException or TimeoutException or System.Net.Sockets.SocketException or SmtpProtocolException => true,
        _ => false
    };

    public async Task<IReadOnlyList<JobUpdate>> RunAsync(MailSettings settings, IReadOnlyList<MailJob> jobs,
        string journalPath, IProgress<JobUpdate>? progress = null, CancellationToken stop = default, bool retryOnly = false)
    {
        BatchPlanner.Validate(settings, true);
        using var journalLock = Storage.AcquireJournal(journalPath);
        var journal = Storage.LoadJournal(journalPath);
        var queue = jobs.Where(job => retryOnly ? job.Status == JobStates.Failed : JobStates.CanSend(job.Status)).ToArray();
        var updates = new List<JobUpdate>();
        foreach (var job in queue)
        {
            if (stop.IsCancellationRequested) break;
            var state = JobStates.Failed;
            var error = "";
            if (journal.Records.TryGetValue(job.Key, out var old) && old.Status is
                (JobStates.Success or JobStates.Sending or JobStates.Uncertain))
            {
                state = old.Status == JobStates.Sending ? JobStates.Uncertain : old.Status;
                error = old.Status == JobStates.Sending ? "上次提交中断，请核实已发送邮件。" : old.Error;
            }
            else
            {
                try
                {
                    if (BatchPlanner.HashFile(job.Filepath) != job.Digest)
                        throw new InvalidDataException("附件在预检后发生变化，请重新检查文件。");
                    journal.Records[job.Key] = JobRecord.From(job) with { Status = JobStates.Sending, Error = "" };
                    // 持久化错误必须直接传播，不能吞掉后继续发送。
                    Storage.WriteAtomic(journalPath, journal);
                    for (var attempt = 0; attempt < settings.MaxRetries; attempt++)
                    {
                        if (stop.IsCancellationRequested) { error = "任务已停止，可稍后继续。"; break; }
                        try
                        {
                            await transport.SendAsync(settings, job).ConfigureAwait(false);
                            state = JobStates.Success;
                            error = "";
                            break;
                        }
                        catch (DeliveryUncertainException exception)
                        { state = JobStates.Uncertain; error = exception.Message; break; }
                        catch (Exception exception)
                        {
                            error = exception.Message;
                            if (attempt + 1 == settings.MaxRetries || !IsTransient(exception)) break;
                            try { await Task.Delay(TimeSpan.FromSeconds(settings.RetryInterval), stop).ConfigureAwait(false); }
                            catch (OperationCanceledException) { error = "任务已停止，可稍后重试。"; break; }
                        }
                    }
                }
                catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException or InvalidDataException)
                { error = exception.Message; }
                journal.Records[job.Key] = JobRecord.From(job) with { Status = state, Error = error };
                Storage.WriteAtomic(journalPath, journal);
            }
            var update = new JobUpdate(job.Key, state, error, updates.Count + 1, queue.Length);
            updates.Add(update);
            progress?.Report(update);
            if (updates.Count < queue.Length)
                try { await Task.Delay(TimeSpan.FromSeconds(settings.SendInterval), stop).ConfigureAwait(false); }
                catch (OperationCanceledException) { break; }
        }
        return updates;
    }
}
