using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Imaging;
using PdfMailer.Core;
using Windows.Data.Pdf;
using Windows.Security.Credentials;
using Windows.Storage.Streams;
using System.Runtime.InteropServices.WindowsRuntime;

namespace PdfMailer.App;

public partial class MainWindow
{
    private string? smokeLibraryDirectory;
    private WorkspaceStore Library => new(smokeLibraryDirectory ?? dataDirectory);
    private string taskId = Guid.NewGuid().ToString("N");
    private string taskName = "";
    private Func<ContentDialog, Task<ContentDialogResult>>? smokeDialogAction;

    private void GroupChanged(object sender, RoutedEventArgs args) => InvalidateIfChanged();
    private void ScopeChanged(object sender, SelectionChangedEventArgs args)
    { if (ScopeInput is not null) ViewModel.SetScope(ScopeInput.SelectedIndex); }
    private void IncludeVisible(object sender, RoutedEventArgs args) { if (!ViewModel.IsBusy) ViewModel.IncludeVisible(true); }
    private void ExcludeVisible(object sender, RoutedEventArgs args) { if (!ViewModel.IsBusy) ViewModel.IncludeVisible(false); }

    private async Task<ContentDialogResult> ContentDialogAsync(string title, object content, string primary, string secondary = "", string close = "取消")
    {
        if (showingDialog) return ContentDialogResult.None;
        showingDialog = true;
        var dialog = new ContentDialog { XamlRoot = Root.XamlRoot, Title = title, Content = content,
            PrimaryButtonText = primary, SecondaryButtonText = secondary, CloseButtonText = close,
            DefaultButton = ContentDialogButton.Close, RequestedTheme = Root.ActualTheme };
        try
        {
            var operation = dialog.ShowAsync();
            ContentDialogResult? simulated = null;
            if (smokeDialogAction is { } action)
            {
                await Task.Delay(250); dialog.UpdateLayout();
                try { simulated = await action(dialog); }
                finally { dialog.Hide(); }
            }
            var result = await operation;
            return simulated ?? result;
        }
        finally { dialog.Content = null; showingDialog = false; }
    }

    private static StackPanel Panel(params UIElement[] children)
    {
        var panel = new StackPanel { Spacing = 12 };
        foreach (var child in children) panel.Children.Add(child);
        return panel;
    }
    private static TextBlock Note(string text) => new() { Text = text, TextWrapping = TextWrapping.Wrap };

    private async void OpenExcelWizard(object sender, RoutedEventArgs args)
    {
        if (IsDemo || ViewModel.IsBusy) return;
        try
        {
            var path = ExcelInput.Text;
            if (!File.Exists(path)) path = await PickFile(".xlsx") ?? "";
            if (path.Length > 0) await ExcelWizard(path);
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async Task ExcelWizard(string path, bool smoke = false)
    {
        if ((IsDemo && !smoke) || ViewModel.IsBusy) return;
        IReadOnlyList<WorksheetPreview> sheets;
        ViewModel.IsChecking = true; ViewModel.IsBusy = true; ViewModel.Notice = "正在读取工作表与示例行…";
        try { sheets = await Task.Run(() => ExcelImport.Inspect(path)); }
        finally { ViewModel.IsChecking = false; ViewModel.IsBusy = false; }
        var sheetInput = new ComboBox { Header = "工作表", ItemsSource = sheets.Select(s => s.Name).ToArray(), HorizontalAlignment = HorizontalAlignment.Stretch };
        var id = new ComboBox { Header = "学号列", HorizontalAlignment = HorizontalAlignment.Stretch };
        var name = new ComboBox { Header = "姓名列", HorizontalAlignment = HorizontalAlignment.Stretch };
        var email = new ComboBox { Header = "邮箱列", HorizontalAlignment = HorizontalAlignment.Stretch };
        var preview = new TextBlock { FontFamily = new Microsoft.UI.Xaml.Media.FontFamily("Consolas"), IsTextSelectionEnabled = true };
        var mapped = Note("");
        void UpdateMapped()
        {
            if (sheetInput.SelectedIndex < 0) return;
            var sheet = sheets[sheetInput.SelectedIndex];
            var columns = new[] { id.SelectedIndex, name.SelectedIndex, email.SelectedIndex };
            mapped.Text = columns.Any(c => c < 0 || c >= sheet.Headers.Count) ? "请为三项分别选择列。" : "映射预览（学号 · 姓名 · 邮箱）\n" +
                string.Join("\n", sheet.Rows.Select(r => string.Join(" · ", columns.Select(c => r[c]))));
        }
        sheetInput.SelectionChanged += (_, _) =>
        {
            var sheet = sheets[sheetInput.SelectedIndex];
            var labels = sheet.Headers.Select((h, i) => $"{i + 1} · {h}").ToArray();
            id.ItemsSource = labels; name.ItemsSource = labels; email.ItemsSource = labels;
            id.SelectedIndex = sheet.IdColumn - 1; name.SelectedIndex = sheet.NameColumn - 1; email.SelectedIndex = sheet.EmailColumn - 1;
            preview.Text = string.Join(" | ", labels) + "\n" + string.Join("\n", sheet.Rows.Select(r => string.Join(" | ", r)));
            UpdateMapped();
        };
        foreach (var box in new[] { id, name, email }) box.SelectionChanged += (_, _) => UpdateMapped();
        sheetInput.SelectedIndex = Math.Max(0, sheets.ToList().FindIndex(s => s.Name == settings.ExcelWorksheet));
        var scroll = new ScrollViewer { Content = preview, MaxHeight = 150, HorizontalScrollBarVisibility = ScrollBarVisibility.Auto };
        var content = new ScrollViewer { MaxHeight = Math.Min(520, Math.Max(220, Root.ActualHeight - 260)), Content = Panel(sheetInput, scroll, id, name, email, mapped,
            Note("第一行为表头。含前导零的学号请在 Excel 中存为文本。")) };
        while (await ContentDialogAsync("导入学生表格", content, "应用映射") == ContentDialogResult.Primary)
        {
            var columns = new[] { id.SelectedIndex, name.SelectedIndex, email.SelectedIndex };
            if (columns.Any(c => c < 0) || columns.Distinct().Count() != 3)
            { mapped.Text = "请选择三个不同的有效列。"; continue; }
            settings.ExcelWorksheet = sheets[sheetInput.SelectedIndex].Name;
            ExcelInput.Text = Path.GetFullPath(path);
            IdColumnInput.Text = (id.SelectedIndex + 1).ToString(); NameColumnInput.Text = (name.SelectedIndex + 1).ToString();
            EmailColumnInput.Text = (email.SelectedIndex + 1).ToString();
            WorksheetLabel.Text = "工作表：" + settings.ExcelWorksheet;
            ViewModel.Invalidate(); ViewModel.Notice = "已应用工作表与列映射，请检查文件。";
            return;
        }
    }

    private async void CheckConnection(object sender, RoutedEventArgs args) => await Diagnose(false);
    private async void CheckAuthentication(object sender, RoutedEventArgs args) => await Diagnose(true);
    private async Task Diagnose(bool authenticate)
    {
        if (IsDemo || ViewModel.IsBusy) return;
        try
        {
            var value = Collect(); ViewModel.IsChecking = true; ViewModel.IsBusy = true;
            ViewModel.Notice = authenticate ? "正在验证登录…" : "正在检测加密连接…";
            await Task.Run(() => new AccountDiagnostics().CheckAsync(value, authenticate));
            await ShowDialog("检测完成", authenticate ? "邮箱认证成功，可以试发给自己。" : "加密连接正常，下一步可验证登录。", "确定");
        }
        catch (Exception exception) { await ShowError(exception.Message); }
        finally { ViewModel.IsChecking = false; ViewModel.IsBusy = false; }
    }

    private async void SendTestClicked(object sender, RoutedEventArgs args)
    {
        if (IsDemo || ViewModel.IsBusy) return;
        try
        {
            if (ViewModel.Snapshot is not { } snapshot || ViewModel.SelectedJob is not { } original || original.Status == JobStates.Blocked)
                throw new InvalidOperationException("请先检查文件并选择一封有效邮件。");
            BatchPlanner.Validate(snapshot, true);
            var result = await ShowDialog("试发给自己", $"实际收件人：{snapshot.SenderEmail}\n内容样本：{original.Name} / {original.StudentId}\n附件：{original.AttachmentCount} 个\n\n主题添加 [试发]，正文与附件保持样本内容。试发单独记录，正式发送状态不受影响。", "确认试发", "取消");
            if (result != ContentDialogResult.Primary) return;
            ViewModel.IsChecking = true; ViewModel.IsBusy = true;
            var test = AccountDiagnostics.CreateTest(original, snapshot.SenderEmail);
            var updates = await Task.Run(() => new BatchSender().RunAsync(snapshot with { MaxRetries = 1 }, [test], Path.Combine(dataDirectory, "test-status.json")));
            var update = updates.Single();
            await ShowDialog("试发结果", update.Status == JobStates.Success ? "SMTP 已接受试发邮件，请到自己的邮箱核对正文与附件。" :
                $"{JobStates.Label(update.Status)}：{update.Error}\n试发记录已单独保存，请先核对邮箱再决定是否再次试发。", "确定");
        }
        catch (Exception exception) { await ShowError(exception.Message); }
        finally { ViewModel.IsChecking = false; ViewModel.IsBusy = false; }
    }

    private async void OpenTemplates(object sender, RoutedEventArgs args) => await EditTemplatesAsync();
    private async Task EditTemplatesAsync(bool smoke = false)
    {
        if ((IsDemo && !smoke) || ViewModel.IsBusy) return;
        try
        {
            var templates = Library.Templates;
            var choose = new ComboBox { Header = "已保存模板", ItemsSource = templates.Select(t => t.Name).ToArray(), HorizontalAlignment = HorizontalAlignment.Stretch };
            var title = new TextBox { Header = "模板名称" };
            var subject = new TextBox { Header = "主题" };
            var body = new TextBox { Header = "正文", AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, Height = 170 };
            var preview = Note("");
            var newButton = new Button { Content = "新建模板" };
            string id = Guid.NewGuid().ToString("N");
            choose.SelectionChanged += (_, _) =>
            {
                if (choose.SelectedIndex < 0) return;
                var selected = templates[choose.SelectedIndex]; id = selected.Id; title.Text = selected.Name; subject.Text = selected.Subject; body.Text = selected.Body;
            };
            newButton.Click += (_, _) => { choose.SelectedIndex = -1; id = Guid.NewGuid().ToString("N"); title.Text = ""; subject.Text = ""; body.Text = ""; };
            void Preview()
            {
                var sample = ViewModel.SelectedJob;
                preview.Text = "实时预览\n" + BatchPlanner.Personalize(subject.Text + "\n\n" + body.Text, sample?.Name ?? "张三", sample?.StudentId ?? "001");
            }
            subject.TextChanged += (_, _) => Preview(); body.TextChanged += (_, _) => Preview();
            var variables = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            foreach (var token in new[] { "{姓名}", "{学号}" })
            {
                var button = new Button { Content = "插入 " + token };
                button.Click += (_, _) => { var target = templateSubjectFocused ? subject : body; target.SelectedText = token; target.Focus(FocusState.Programmatic); };
                variables.Children.Add(button);
            }
            subject.GotFocus += (_, _) => templateSubjectFocused = true;
            body.GotFocus += (_, _) => templateSubjectFocused = false;
            if (File.Exists(TemplateInput.Text))
            {
                var current = BatchPlanner.LoadTemplate(TemplateInput.Text);
                title.Text = Path.GetFileNameWithoutExtension(TemplateInput.Text); subject.Text = current.Subject; body.Text = current.Body;
            }
            Preview();
            var content = new ScrollViewer { MaxHeight = Math.Min(560, Math.Max(220, Root.ActualHeight - 260)), Content = Panel(choose, newButton, title, subject, variables, body,
                new ScrollViewer { Content = preview, MaxHeight = 140 }) };
            while (true)
            {
                var result = await ContentDialogAsync("邮件模板库", content, "保存并应用", "仅保存到模板库");
                if (result == ContentDialogResult.None) return;
                var template = new SavedTemplate(id, title.Text.Trim(), subject.Text.Trim(), body.Text.Trim());
                try { Library.SaveTemplate(template); }
                catch (InvalidDataException exception) { preview.Text = exception.Message; continue; }
                if (result == ContentDialogResult.Primary) { TemplateInput.Text = Library.Materialize(template); ViewModel.Invalidate(); }
                ViewModel.Notice = "模板已保存。"; return;
            }
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }
    private bool templateSubjectFocused;
    private string? rememberContext;
    private string RememberContext => $"{SenderInput?.Text.Trim().ToLowerInvariant()}|{ServerInput?.Text.Trim().ToLowerInvariant()}|{PortInput?.Value}|{SecurityInput?.SelectedIndex}";
    private void RememberChanged(object sender, RoutedEventArgs args) => rememberContext = RememberInput.IsChecked == true ? RememberContext : null;
    private void ValidateRememberContext()
    {
        if (loaded && RememberInput?.IsChecked == true && rememberContext != RememberContext) RememberInput.IsChecked = false;
    }

    private void SaveCurrentTask(MailSettings value)
    {
        Library.SaveTask(new(taskId, taskName.Length > 0 ? taskName : Path.GetFileName(value.AttachmentFolder.TrimEnd(Path.DirectorySeparatorChar)),
            DateTimeOffset.Now, value, ViewModel.TotalCount, ViewModel.SuccessCount, ViewModel.AttentionCount));
    }
    private async void SaveTaskClicked(object sender, RoutedEventArgs args)
    {
        if (IsDemo || ViewModel.IsBusy) return;
        try
        {
            var name = new TextBox { Header = "任务名称", Text = taskName, PlaceholderText = "例如：秋季成绩单" };
            if (await ContentDialogAsync("保存任务", name, "保存") != ContentDialogResult.Primary) return;
            if (string.IsNullOrWhiteSpace(name.Text)) throw new InvalidDataException("请输入任务名称。");
            taskName = name.Text.Trim(); SaveCurrentTask(Collect()); ViewModel.Notice = "任务已保存，可在任务历史中恢复。";
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }
    private async void OpenTasks(object sender, RoutedEventArgs args) => await ManageTasksAsync();
    private async Task ManageTasksAsync(bool smoke = false)
    {
        if ((IsDemo && !smoke) || ViewModel.IsBusy) return;
        try
        {
            var tasks = Library.Tasks;
            var list = new ListView { MaxHeight = 330, SelectionMode = ListViewSelectionMode.Single,
                ItemsSource = tasks.Select(t => $"{t.Name}\n{t.UpdatedAt.LocalDateTime:g} · {t.Total} 封 / 成功 {t.Success} / 需关注 {t.Attention}").ToArray() };
            var result = await ContentDialogAsync("任务历史", Panel(list, Note("恢复或复制均保留原发送日志，重新检查后会跳过已发送项。复制后可修改文件、模板和账户。")), "恢复任务", "复制为新任务");
            if (result == ContentDialogResult.None || list.SelectedIndex < 0) return;
            var selected = tasks[list.SelectedIndex];
            taskId = result == ContentDialogResult.Primary ? selected.Id : Guid.NewGuid().ToString("N");
            taskName = result == ContentDialogResult.Primary ? selected.Name : selected.Name + "（副本）";
            settings = selected.Settings with { SenderPassword = "" };
            FillForm(settings, settings.StateFile); RestoreRememberedPassword(); ViewModel.Invalidate();
            ViewModel.Notice = "已载入任务，请重新检查文件与发送日志。";
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private static string CredentialResource(SavedAccount account) => "PdfMailer/" + account.Server.ToLowerInvariant() + "/" + account.Port + "/" + account.UseSsl;
    private void RestoreRememberedPassword()
    {
        var account = Library.Accounts.FirstOrDefault(a => a.Email.Equals(SenderInput.Text.Trim(), StringComparison.OrdinalIgnoreCase)
            && a.Server.Equals(ServerInput.Text.Trim(), StringComparison.OrdinalIgnoreCase) && a.Port == PortInput.Value && a.UseSsl == (SecurityInput.SelectedIndex == 0));
        RememberInput.IsChecked = account?.RememberPassword == true;
        if (account?.RememberPassword != true) return;
        try { var credential = new PasswordVault().Retrieve(CredentialResource(account), account.Email); credential.RetrievePassword(); PasswordInput.Password = credential.Password; }
        catch (System.Runtime.InteropServices.COMException) { ViewModel.Notice = "已保存账户的授权码不可用，请重新输入并保存账户。"; }
    }
    private async void SaveAccountClicked(object sender, RoutedEventArgs args)
    {
        if (IsDemo || ViewModel.IsBusy) return;
        try
        {
            var account = PersistAccount(Collect());
            await ShowDialog("账户已保存", account.RememberPassword ? "授权码已交给 Windows 凭据保管库。" : "已保存账户信息，授权码仅在本次会话使用。", "确定");
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }
    private SavedAccount PersistAccount(MailSettings value)
    {
        ValidateRememberContext();
        if (!BatchPlanner.ValidEmail(value.SenderEmail)) throw new InvalidDataException("请输入有效的发件邮箱。");
        var account = new SavedAccount(value.SenderEmail.ToLowerInvariant(), value.SmtpServer, value.SmtpPort, value.UseSsl, RememberInput.IsChecked == true);
        if (account.RememberPassword && value.SenderPassword.Length == 0) throw new InvalidDataException("请输入要安全保存的授权码。");
        var old = Library.Accounts.FirstOrDefault(a => a.Email.Equals(account.Email, StringComparison.OrdinalIgnoreCase));
        if (account.RememberPassword) new PasswordVault().Add(new PasswordCredential(CredentialResource(account), account.Email, value.SenderPassword));
        if (old is not null && (!account.RememberPassword || CredentialResource(old) != CredentialResource(account))) ForgetCredential(old);
        Library.SaveAccount(account);
        return account;
    }
    private static void ForgetCredential(SavedAccount account)
    {
        var vault = new PasswordVault();
        try { vault.Remove(vault.Retrieve(CredentialResource(account), account.Email)); }
        catch (System.Runtime.InteropServices.COMException exception) when (exception.HResult == unchecked((int)0x80070490)) { }
    }
    private async void OpenAccounts(object sender, RoutedEventArgs args)
    {
        if (IsDemo || ViewModel.IsBusy) return;
        try
        {
            var accounts = Library.Accounts;
            var list = new ListView { ItemsSource = accounts.Select(a => $"{a.Email}\n{a.Server}:{a.Port} · " + (a.RememberPassword ? "已安全记住授权码" : "每次输入授权码")).ToArray(), MaxHeight = 350 };
            var result = await ContentDialogAsync("发件账户", list, "使用账户", "删除账户");
            if (result == ContentDialogResult.None || list.SelectedIndex < 0) return;
            var account = accounts[list.SelectedIndex];
            if (result == ContentDialogResult.Secondary)
            {
                ForgetCredential(account); Library.DeleteAccount(account.Email); ViewModel.Notice = "账户和保存的授权码已移除。"; return;
            }
            SenderInput.Text = account.Email; ServerInput.Text = account.Server; PortInput.Value = account.Port;
            SecurityInput.SelectedIndex = account.UseSsl ? 0 : 1; PasswordInput.Password = "";
            RestoreRememberedPassword(); ViewModel.Invalidate();
        }
        catch (Exception exception) { await ShowError(exception.Message); }
    }

    private async void OpenPdfPreview(object sender, RoutedEventArgs args)
    {
        if (ViewModel.SelectedJob is not { } job) return;
        try { await PdfPreviewAsync(job); }
        catch (Exception exception) { await ShowError("无法预览 PDF：" + exception.Message); }
    }
    private async Task PdfPreviewAsync(MailJob job)
    {
        var files = job.Files;
        var fileInput = new ComboBox { ItemsSource = files.Select(f => f.Filename).ToArray(), HorizontalAlignment = HorizontalAlignment.Stretch };
        var pageInput = new NumberBox { Header = "页码", Minimum = 1, Maximum = 1, Value = 1, SpinButtonPlacementMode = NumberBoxSpinButtonPlacementMode.Compact, Width = 100 };
        var zoom = new ComboBox { ItemsSource = new[] { "适合宽度", "100%", "150%", "200%" }, SelectedIndex = 0, Width = 130 };
        var image = new Image { Stretch = Microsoft.UI.Xaml.Media.Stretch.Uniform, HorizontalAlignment = HorizontalAlignment.Left };
        var notice = Note("");
        var previous = new Button { Content = "上一页" }; var next = new Button { Content = "下一页" };
        PdfDocument? document = null; int version = 0; bool adjusting = false; bool closed = false;
        var sources = new List<IRandomAccessStream>();
        async Task Render()
        {
            if (adjusting || closed || document is null) return;
            var token = ++version;
            try
            {
                var number = Math.Clamp((int)(double.IsFinite(pageInput.Value) ? pageInput.Value : 1), 1, (int)document.PageCount);
                using var page = document.GetPage((uint)(number - 1));
                using var stream = new InMemoryRandomAccessStream();
                var width = zoom.SelectedIndex == 0 ? 450d : page.Size.Width * (zoom.SelectedIndex == 1 ? 1 : zoom.SelectedIndex == 2 ? 1.5 : 2);
                await page.RenderToStreamAsync(stream, new PdfPageRenderOptions { DestinationWidth = (uint)Math.Min(2400, width * 1.5) });
                var bitmap = new BitmapImage(); stream.Seek(0); await bitmap.SetSourceAsync(stream);
                if (token != version || closed) return;
                image.Source = bitmap; image.Width = width;
                previous.IsEnabled = number > 1; next.IsEnabled = number < document.PageCount;
                notice.Text = $"第 {number} / {document.PageCount} 页";
            }
            catch (Exception exception) { if (token == version && !closed) { image.Source = null; notice.Text = "无法显示该页：" + exception.Message; } }
        }
        async Task Load()
        {
            if (closed || fileInput.SelectedIndex < 0) return;
            var token = ++version; document = null; image.Source = null; notice.Text = "正在加载 PDF…";
            try
            {
                var bytes = await File.ReadAllBytesAsync(Path.GetFullPath(files[fileInput.SelectedIndex].Filepath));
                if (closed || token != version) return;
                var source = new InMemoryRandomAccessStream(); sources.Add(source);
                await source.WriteAsync(bytes.AsBuffer()); source.Seek(0);
                var loadedDocument = await PdfDocument.LoadFromStreamAsync(source);
                if (token != version || closed) return;
                document = loadedDocument;
                adjusting = true; pageInput.Maximum = document.PageCount; pageInput.Value = 1; adjusting = false;
                await Render();
            }
            catch (Exception exception) { if (token == version && !closed) notice.Text = "PDF 无法打开（可能已损坏、加密或被移动）：" + exception.Message; }
        }
        fileInput.SelectionChanged += async (_, _) => await Load();
        pageInput.ValueChanged += async (_, _) => await Render(); zoom.SelectionChanged += async (_, _) => await Render();
        previous.Click += (_, _) => pageInput.Value = Math.Max(1, pageInput.Value - 1);
        next.Click += (_, _) => pageInput.Value = Math.Min(pageInput.Maximum, pageInput.Value + 1);
        var tools = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        foreach (var element in new UIElement[] { previous, pageInput, next, zoom }) tools.Children.Add(element);
        fileInput.SelectedIndex = 0;
        try { await ContentDialogAsync("PDF 附件预览", Panel(fileInput, tools, notice, new ScrollViewer { Content = image, Height = Math.Min(440, Math.Max(160, Root.ActualHeight - 340)),
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollMode = ScrollMode.Enabled }), "", close: "关闭"); }
        finally
        {
            closed = true; version++; document = null; image.Source = null;
            foreach (var source in sources) source.Dispose();
        }
    }
}
