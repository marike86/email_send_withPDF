using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using PdfMailer.Core;

namespace PdfMailer.App;

public sealed partial class MailerViewModel : ObservableObject
{
    public ObservableCollection<MailJob> VisibleJobs { get; } = [];
    public IReadOnlyList<MailJob> Jobs { get; private set; } = [];
    public MailSettings? Snapshot { get; private set; }
    private CancellationTokenSource? stop;
    private string search = "";
    private int filter;
    [ObservableProperty] public partial bool IsBusy { get; set; }
    [ObservableProperty] public partial bool CanSend { get; set; }
    [ObservableProperty] public partial bool CanRetry { get; set; }
    [ObservableProperty] public partial bool CanExport { get; set; }
    [ObservableProperty] public partial bool CanResolve { get; set; }
    [ObservableProperty] public partial bool IsChecking { get; set; }
    [ObservableProperty] public partial int TotalCount { get; set; }
    [ObservableProperty] public partial int PendingCount { get; set; }
    [ObservableProperty] public partial int SuccessCount { get; set; }
    [ObservableProperty] public partial int AttentionCount { get; set; }
    [ObservableProperty] public partial double Progress { get; set; }
    [ObservableProperty] public partial string Notice { get; set; } = "选择文件并填写发件账户，然后检查匹配结果。预检不会连接邮箱。";
    [ObservableProperty] public partial MailJob? SelectedJob { get; set; }
    public bool CanEdit => !IsBusy;
    public bool CanStop => IsBusy && !IsChecking;

    partial void OnIsBusyChanged(bool value)
    {
        OnPropertyChanged(nameof(CanEdit)); OnPropertyChanged(nameof(CanStop)); UpdateActions();
    }
    partial void OnIsCheckingChanged(bool value) => OnPropertyChanged(nameof(CanStop));
    partial void OnSelectedJobChanged(MailJob? value) => UpdateActions();

    public void Invalidate()
    {
        Snapshot = null;
        Notice = "设置已更改，请重新检查文件。";
        UpdateActions();
    }

    private void UpdateActions()
    {
        CanSend = !IsBusy && Snapshot is not null && Jobs.Any(j => JobStates.CanSend(j.Status));
        CanRetry = !IsBusy && Snapshot is not null && Jobs.Any(j => j.Status == JobStates.Failed);
        CanExport = !IsBusy && Jobs.Count > 0;
        CanResolve = !IsBusy && Snapshot is not null && SelectedJob?.Status is JobStates.Sending or JobStates.Uncertain;
    }

    private void RefreshCounts()
    {
        TotalCount = Jobs.Count;
        PendingCount = Jobs.Count(j => JobStates.CanSend(j.Status));
        SuccessCount = Jobs.Count(j => j.Status == JobStates.Success);
        AttentionCount = Jobs.Count(j => JobStates.NeedsAttention(j.Status));
        UpdateActions();
    }

    public void Filter(string term, int selection)
    {
        search = term; filter = selection;
        var selected = SelectedJob;
        VisibleJobs.Clear();
        foreach (var job in Jobs)
        {
            var match = $"{job.Filename} {job.StudentId} {job.Name} {job.Email}".Contains(term.Trim(), StringComparison.OrdinalIgnoreCase);
            var state = selection switch { 1 => JobStates.CanSend(job.Status), 2 => job.Status == JobStates.Success,
                3 => JobStates.NeedsAttention(job.Status), _ => true };
            if (match && state) VisibleJobs.Add(job);
        }
        SelectedJob = selected is not null && VisibleJobs.Contains(selected) ? selected : VisibleJobs.FirstOrDefault();
    }

    public async Task CheckAsync(MailSettings settings, string statePath)
    {
        if (IsBusy) return;
        Snapshot = null;
        IsChecking = true; IsBusy = true; Progress = 0;
        Notice = "正在检查 PDF、Excel 和发送日志…";
        try
        {
            Jobs = await Task.Run(() => BatchPlanner.Prepare(settings, statePath));
            Snapshot = settings with { };
            Filter(search, filter); RefreshCounts();
            Notice = TotalCount == 0 ? "该目录未找到 PDF 文件。" : $"检查完成：{TotalCount} 个 PDF，{AttentionCount} 项需关注。选择一行预览邮件。";
        }
        catch { Invalidate(); throw; }
        finally { IsChecking = false; IsBusy = false; }
    }

    public async Task SendAsync(string statePath, bool retryOnly)
    {
        if (IsBusy || Snapshot is null) return;
        var settings = Snapshot with { };
        var jobs = Jobs.ToArray();
        stop = new CancellationTokenSource();
        IsBusy = true; Progress = 0;
        Notice = "正在发送，可停止后续邮件；当前网络请求完成后停止。";
        try
        {
            // Progress 在 UI 线程创建，后台不会直接修改绑定对象。
            var progress = new Progress<JobUpdate>(Apply);
            var results = await Task.Run(() => new BatchSender().RunAsync(settings, jobs, statePath, progress, stop.Token, retryOnly));
            foreach (var result in results) Apply(result);
            Filter(search, filter);
            Notice = stop.IsCancellationRequested ? "已停止后续发送，进度已保存，可稍后继续。" : "本轮处理完成，可以导出结果。";
        }
        catch { Invalidate(); throw; }
        finally { stop.Dispose(); stop = null; IsBusy = false; }
    }

    private void Apply(JobUpdate update)
    {
        var job = Jobs.FirstOrDefault(j => j.Key == update.Key);
        if (job is null) return;
        job.Status = update.Status; job.Error = update.Error;
        Progress = update.Total == 0 ? 0 : 100d * update.Completed / update.Total;
        Notice = $"{update.Completed} / {update.Total} · {job.StatusLabel} · {job.Filename}";
        RefreshCounts();
    }

    public void Stop()
    {
        stop?.Cancel();
        Notice = "正在停止后续发送，请等待当前网络请求完成。";
    }

    public void Resolve(string statePath, bool delivered)
    {
        if (SelectedJob is not { } job || !CanResolve) return;
        Storage.ResolveUncertain(statePath, job.Key, delivered);
        job.Status = delivered ? JobStates.Success : JobStates.Failed;
        job.Error = "用户已核实发送结果";
        RefreshCounts();
    }

    public void LoadDemo(IReadOnlyList<MailJob> jobs, MailSettings? settings = null)
    {
        Jobs = jobs; Snapshot = settings ?? new(); Filter("", 0); RefreshCounts();
        Notice = "演示数据 · 不会连接 SMTP 或发送邮件。";
    }
}
