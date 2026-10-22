using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using System.Runtime.InteropServices;
using PdfMailer.Core;
using Windows.Storage.Pickers;

namespace PdfMailer.App;

public partial class MainWindow : Window
{
    public MailerViewModel ViewModel { get; } = new();
    private readonly string dataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PdfMailer");
    private string configPath = "";
    private MailSettings settings = new();
    private bool loaded;
    private bool showingDialog;
    private string? startupError;
    private readonly string[] arguments = Environment.GetCommandLineArgs();
    private bool IsDemo => arguments.Contains("--demo") || arguments.Contains("--smoke-test");

    public MainWindow()
    {
        InitializeComponent();
        Root.DataContext = ViewModel;
        Root.SizeChanged += (_, _) =>
        {
            // 矮窗口优先保留邮件阅读区域，统计摘要仍可从发送结果中查看。
            var compact = Root.ActualHeight < 780;
            MetricsPanel.Visibility = compact ? Visibility.Collapsed : Visibility.Visible;
            WorkspaceSubtitle.Visibility = compact ? Visibility.Collapsed : Visibility.Visible;
        };
        if (Navigation.SettingsItem is NavigationViewItem settingsItem) settingsItem.Content = "账户与设置";
        ViewModel.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(ViewModel.SelectedJob))
                DispatcherQueue.TryEnqueue(() => PreviewScroll.ChangeView(null, 0, null, true));
        };
        Title = "PDF 邮件工作台";
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        SystemBackdrop = new MicaBackdrop();
        var scale = GetDpiForWindow(WinRT.Interop.WindowNative.GetWindowHandle(this)) / 96d;
        AppWindow.Resize(new Windows.Graphics.SizeInt32((int)(1320 * scale), (int)(940 * scale)));
        AppWindow.Closing += (_, args) =>
        {
            if (!ViewModel.IsBusy) return;
            args.Cancel = true;
            ViewModel.Stop();
            ViewModel.Notice = "请等待当前任务停止后再关闭窗口。";
        };
        configPath = GetArgument("--config") ?? Path.Combine(dataDirectory, "config.json");
        try { settings = Storage.LoadSettings(configPath); }
        catch (Exception exception) { startupError = exception.Message; }
        FillForm(settings, GetArgument("--state") ?? (settings.StateFile.Length > 0 ? settings.StateFile : Path.Combine(dataDirectory, "status.json")));
        loaded = true;
        Root.Loaded += RootLoaded;
    }

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr window);

    private string? GetArgument(string flag)
    {
        var index = Array.IndexOf(arguments, flag);
        return index >= 0 && index + 1 < arguments.Length ? arguments[index + 1] : null;
    }

    private async void RootLoaded(object sender, RoutedEventArgs args)
    {
        Root.Loaded -= RootLoaded;
        if (Navigation.SettingsItem is NavigationViewItem settingsItem) settingsItem.Content = "账户与设置";
        if (IsDemo) LoadDemo();
        if (startupError is not null && !IsDemo) await ShowError(startupError);
        if (GetArgument("--smoke-test") is { } output)
            await RunSmokeAsync(output);
    }

    private void FillForm(MailSettings value, string journal)
    {
        var previous = loaded; loaded = false;
        PdfFolderInput.Text = value.AttachmentFolder;
        ExcelInput.Text = value.StudentInfoExcel;
        TemplateInput.Text = value.EmailTemplate;
        SenderInput.Text = value.SenderEmail;
        PasswordInput.Password = value.SenderPassword;
        ServerInput.Text = value.SmtpServer;
        PortInput.Value = value.SmtpPort;
        SecurityInput.SelectedIndex = value.UseSsl || value.SmtpPort == 465 ? 0 : 1;
        IdColumnInput.Text = value.ExcelStudentIdCol.ToString();
        NameColumnInput.Text = value.ExcelNameCol.ToString();
        EmailColumnInput.Text = value.ExcelEmailCol.ToString();
        SendIntervalInput.Value = value.SendInterval;
        RetriesInput.Value = value.MaxRetries;
        RetryIntervalInput.Value = value.RetryInterval;
        JournalInput.Text = journal;
        ConfigLocation.Text = "配置文件：" + Path.GetFullPath(configPath);
        loaded = previous;
    }

    private MailSettings Collect()
    {
        static int Integer(NumberBox control, string title)
        {
            var value = control.Value;
            if (!double.IsFinite(value) || value != Math.Truncate(value))
                throw new InvalidDataException(title + "应为整数。");
            return checked((int)value);
        }
        var directory = Path.GetDirectoryName(Path.GetFullPath(configPath))!;
        var value = settings with
        {
            SenderEmail = SenderInput.Text.Trim(), SenderPassword = PasswordInput.Password,
            SmtpServer = ServerInput.Text.Trim(), SmtpPort = Integer(PortInput, "端口"),
            UseSsl = SecurityInput.SelectedIndex == 0, UseTls = SecurityInput.SelectedIndex == 1,
            AttachmentFolder = Storage.Resolve(PdfFolderInput.Text, directory),
            StudentInfoExcel = Storage.Resolve(ExcelInput.Text, directory),
            EmailTemplate = Storage.Resolve(TemplateInput.Text, directory),
            StateFile = Storage.Resolve(JournalInput.Text, directory),
            ExcelStudentIdCol = BatchPlanner.ColumnNumber(IdColumnInput.Text),
            ExcelNameCol = BatchPlanner.ColumnNumber(NameColumnInput.Text),
            ExcelEmailCol = BatchPlanner.ColumnNumber(EmailColumnInput.Text),
            SendInterval = SendIntervalInput.Value, MaxRetries = Integer(RetriesInput, "最多尝试次数"),
            RetryInterval = RetryIntervalInput.Value
        };
        BatchPlanner.Validate(value);
        if (value.StateFile.Length == 0) throw new InvalidDataException("请输入发送日志路径。");
        return value;
    }

    private void InvalidateIfChanged()
    {
        if (!loaded || ViewModel.Snapshot is null) return;
        try { if (Collect() != ViewModel.Snapshot) ViewModel.Invalidate(); }
        catch { ViewModel.Invalidate(); }
    }
    private void SettingsTextChanged(object sender, TextChangedEventArgs args) => InvalidateIfChanged();
    private void SettingsPasswordChanged(object sender, RoutedEventArgs args) => InvalidateIfChanged();
    private void SettingsNumberChanged(NumberBox sender, NumberBoxValueChangedEventArgs args) => InvalidateIfChanged();
    private void SettingsSelectionChanged(object sender, SelectionChangedEventArgs args) => InvalidateIfChanged();
    private void FilterChanged(object sender, RoutedEventArgs args)
    { if (loaded) ViewModel.Filter(SearchInput.Text, StateFilter.SelectedIndex); }

    private void NavigationChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (Workspace is null || SettingsPanel is null) return;
        Workspace.Visibility = args.IsSettingsSelected ? Visibility.Collapsed : Visibility.Visible;
        SettingsPanel.Visibility = args.IsSettingsSelected ? Visibility.Visible : Visibility.Collapsed;
    }

    private void InitializePicker(object picker) => WinRT.Interop.InitializeWithWindow.Initialize(picker, WinRT.Interop.WindowNative.GetWindowHandle(this));

    private async void PickFolder(object sender, RoutedEventArgs args)
    {
        try
        {
            var picker = new FolderPicker(); picker.FileTypeFilter.Add("*"); InitializePicker(picker);
            if (await picker.PickSingleFolderAsync() is { } folder) PdfFolderInput.Text = folder.Path;
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async Task<string?> PickFile(params string[] extensions)
    {
        var picker = new FileOpenPicker();
        foreach (var extension in extensions) picker.FileTypeFilter.Add(extension);
        InitializePicker(picker);
        return (await picker.PickSingleFileAsync())?.Path;
    }

    private async void PickExcel(object sender, RoutedEventArgs args)
    { try { if (await PickFile(".xlsx") is { } path) ExcelInput.Text = path; } catch (Exception exception) { await ShowError(exception.Message); } }
    private async void PickTemplate(object sender, RoutedEventArgs args)
    { try { if (await PickFile(".txt") is { } path) TemplateInput.Text = path; } catch (Exception exception) { await ShowError(exception.Message); } }

    private async void CheckFiles(object sender, RoutedEventArgs args)
    {
        if (IsDemo) { ViewModel.Notice = "演示模式不访问文件或邮箱，请正常启动后使用自己的资料。"; return; }
        try
        {
            settings = Collect();
            await ViewModel.CheckAsync(settings, settings.StateFile);
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async void SaveSettings(object sender, RoutedEventArgs args)
    {
        if (IsDemo) { ViewModel.Notice = "演示模式不会保存设置。"; return; }
        try
        {
            settings = Collect();
            Storage.WriteAtomic(configPath, settings);
            await ShowDialog("设置已保存", "已保存非敏感配置，授权码未写入文件。", "确定");
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async void ImportSettings(object sender, RoutedEventArgs args)
    {
        if (IsDemo) return;
        try
        {
            if (await PickFile(".json") is not { } path) return;
            var imported = Storage.LoadSettings(path);
            // 使用导入位置旁的旧日志，避免意外改用空日志重发。
            var adjacentState = Path.Combine(Path.GetDirectoryName(path)!, "status.json");
            var state = imported.StateFile.Length > 0 ? imported.StateFile : File.Exists(adjacentState) ? adjacentState : JournalInput.Text;
            settings = imported;
            FillForm(imported, state);
            ViewModel.Invalidate();
            ViewModel.Notice = "已导入配置，请核对发送日志与文件路径后重新检查。保存时将写入当前应用配置。";
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async void SendEmails(object sender, RoutedEventArgs args) => await Send(false);
    private async void RetryFailed(object sender, RoutedEventArgs args) => await Send(true);
    private async Task Send(bool retryOnly)
    {
        if (IsDemo) { ViewModel.Notice = "演示模式不会连接 SMTP 或发送邮件。"; return; }
        try
        {
            if (ViewModel.Snapshot is not { } snapshot || ViewModel.IsBusy) return;
            BatchPlanner.Validate(snapshot, true);
            var count = ViewModel.Jobs.Count(job => retryOnly ? job.Status == JobStates.Failed : JobStates.CanSend(job.Status));
            var result = await ShowDialog("确认发送计划", $"发件人：{snapshot.SenderEmail}\n将发送 {count} 封带 PDF 的邮件。\n搜索和筛选只影响显示，发送范围为整份计划。\n\n请确认已核对预览内容。", "确认发送", "取消");
            if (result == ContentDialogResult.Primary) await ViewModel.SendAsync(snapshot.StateFile, retryOnly);
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private void StopSending(object sender, RoutedEventArgs args) => ViewModel.Stop();

    private async void ExportResults(object sender, RoutedEventArgs args)
    {
        try
        {
            var picker = new FileSavePicker { SuggestedFileName = "邮件发送结果" };
            picker.FileTypeChoices.Add("Excel 工作簿", new List<string> { ".xlsx" });
            picker.FileTypeChoices.Add("CSV 表格", new List<string> { ".csv" });
            InitializePicker(picker);
            if (await picker.PickSaveFileAsync() is not { } file) return;
            var jobs = ViewModel.Jobs;
            await Task.Run(() => ResultExporter.Export(jobs, file.Path));
            ViewModel.Notice = "结果已导出：" + file.Path;
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async void ResolveUncertain(object sender, RoutedEventArgs args)
    {
        try
        {
            if (IsDemo || ViewModel.SelectedJob is not { } job || ViewModel.Snapshot is null) return;
            var result = await ShowDialog("核实发送结果", $"请先在邮箱中核对：\n{job.Filename}\n{job.Email}\n\n确认是否已经送达后再选择。", "已送达，标记成功", "取消", "确认未送达，允许重试");
            if (result == ContentDialogResult.None) return;
            ViewModel.Resolve(ViewModel.Snapshot.StateFile, result == ContentDialogResult.Primary);
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async Task ShowError(string message)
    {
        ViewModel.Notice = "操作未完成：" + message;
        await ShowDialog("请检查", message, "确定");
    }

    private async Task<ContentDialogResult> ShowDialog(string title, string message, string primary, string close = "", string secondary = "")
    {
        if (showingDialog) return ContentDialogResult.None;
        showingDialog = true;
        try
        {
            var dialog = new ContentDialog
            {
                XamlRoot = Root.XamlRoot, Title = title, Content = new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap },
                PrimaryButtonText = primary, CloseButtonText = close, SecondaryButtonText = secondary,
                DefaultButton = close.Length > 0 ? ContentDialogButton.Close : ContentDialogButton.Primary
            };
            return await dialog.ShowAsync();
        }
        finally { showingDialog = false; }
    }

    private void LoadDemo()
    {
        var demoSettings = new MailSettings
        {
            SenderEmail = "sender@example.com", AttachmentFolder = @"D:\邮件资料\PDF附件",
            StudentInfoExcel = @"D:\邮件资料\学生信息.xlsx", EmailTemplate = @"D:\邮件资料\邮件模板.txt"
        };
        FillForm(demoSettings, @"D:\邮件资料\status.json");
        ViewModel.LoadDemo([
            new() { Filename = "001_学习资料.pdf", StudentId = "001", Name = "张三", Email = "zhangsan@example.com", Subject = "张三同学的学习资料", Body = "张三同学：\n\n您好！学号 001 的学习资料已随邮件以 PDF 附件发送，请查收。\n\n如有疑问，请回复本邮件。\n\n祝学习顺利！", Key = "demo1" },
            new() { Filename = "002_学习资料.pdf", StudentId = "002", Name = "李明", Email = "liming@example.com", Subject = "李明同学的学习资料", Body = "李明同学：\n\n请查收附件。", Key = "demo2" },
            new() { Filename = "003_学习资料.pdf", StudentId = "003", Name = "王悦", Email = "wangyue@example.com", Subject = "王悦同学的学习资料", Body = "王悦同学：\n\n请查收附件。", Key = "demo3", Status = JobStates.Success },
            new() { Filename = "004_学习资料.pdf", StudentId = "004", Status = JobStates.Blocked, Error = "Excel 中找不到此学号。", Key = "demo4" },
            new() { Filename = "待检查.pdf", Status = JobStates.Blocked, Error = "文件名无法提取学号。", Key = "demo5" }
        ], Collect());
    }
}

public sealed class HasTextConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language) => value is string text && !string.IsNullOrWhiteSpace(text);
    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotSupportedException();
}

public sealed class VisibleConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language) => value is true ? Visibility.Visible : Visibility.Collapsed;
    public object ConvertBack(object value, Type targetType, object parameter, string language) => throw new NotSupportedException();
}
