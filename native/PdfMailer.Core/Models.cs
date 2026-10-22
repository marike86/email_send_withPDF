using System.Text.Json.Serialization;
using CommunityToolkit.Mvvm.ComponentModel;

namespace PdfMailer.Core;

public sealed record MailSettings
{
    public string SmtpServer { get; set; } = "smtp.qq.com";
    public int SmtpPort { get; set; } = 465;
    public bool UseSsl { get; set; } = true;
    public bool UseTls { get; set; }
    public string SenderEmail { get; set; } = "";
    [JsonIgnore] public string SenderPassword { get; set; } = "";
    public string AttachmentFolder { get; set; } = "";
    public string StudentInfoExcel { get; set; } = "";
    public string EmailTemplate { get; set; } = "";
    public string StateFile { get; set; } = "";
    public int ExcelStudentIdCol { get; set; } = 1;
    public int ExcelNameCol { get; set; } = 2;
    public int ExcelEmailCol { get; set; } = 3;
    public double SendInterval { get; set; } = 3;
    public int MaxRetries { get; set; } = 3;
    public double RetryInterval { get; set; } = 10;
}

public static class JobStates
{
    public const string Ready = "ready", Success = "success", Failed = "failed",
        Blocked = "blocked", Sending = "sending", Uncertain = "uncertain";
    public static bool CanSend(string state) => state is Ready or Failed;
    public static bool NeedsAttention(string state) => state is Failed or Blocked or Sending or Uncertain;
    public static string Label(string state) => state switch
    {
        Ready => "待发送", Success => "已发送", Failed => "失败", Blocked => "需修正",
        Sending or Uncertain => "待核实", _ => state
    };
}

public sealed partial class MailJob : ObservableObject
{
    public string Filename { get; init; } = "";
    public string Filepath { get; set; } = "";
    public string StudentId { get; init; } = "";
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Body { get; set; } = "";
    public string Digest { get; set; } = "";
    public string Key { get; set; } = "";
    [ObservableProperty, NotifyPropertyChangedFor(nameof(StatusLabel))]
    private string status = JobStates.Ready;
    [ObservableProperty] private string error = "";
    public string StatusLabel => JobStates.Label(Status);
    public string Recipient => $"{Name} <{Email}>";
}

public sealed record Student(string Name, string Email);
public sealed record MailTemplate(string Subject, string Body);
public sealed record JobRecord
{
    public string Filename { get; init; } = "";
    public string StudentId { get; init; } = "";
    public string Name { get; init; } = "";
    public string Email { get; init; } = "";
    public string Status { get; set; } = JobStates.Sending;
    public string Error { get; set; } = "";
    public string Time { get; set; } = DateTimeOffset.Now.ToString("O");
    public static JobRecord From(MailJob job) => new()
    {
        Filename = job.Filename, StudentId = job.StudentId, Name = job.Name, Email = job.Email,
        Status = job.Status, Error = job.Error
    };
}

public sealed record SendJournal
{
    public int Version { get; init; } = 2;
    public Dictionary<string, JobRecord> Records { get; init; } = [];
}

public sealed record JobUpdate(string Key, string Status, string Error, int Completed, int Total);
public sealed class DeliveryUncertainException(string message, Exception? inner = null) : Exception(message, inner);
