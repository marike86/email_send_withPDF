using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using PdfMailer.Core;
using System.Text;

namespace PdfMailer.App;

public partial class MainWindow
{
    private async Task VerifyFeaturesAsync(string output, Dictionary<string, object> checks)
    {
        LoadDemo();
        ScopeInput.SelectedIndex = 1;
        ViewModel.IncludeVisible(false);
        if (ViewModel.CanSend || ViewModel.SelectedCount != 0) throw new InvalidOperationException("清除勾选没有更新发送范围。");
        ViewModel.Jobs[0].IsIncluded = true;
        if (ViewModel.SelectedCount != 1) throw new InvalidOperationException("勾选发送范围错误。");
        ScopeInput.SelectedIndex = 2; SearchInput.Text = "李明";
        await Task.Delay(150);
        if (ViewModel.SendQueue().Count != 1 || ViewModel.SendQueue()[0].Name != "李明") throw new InvalidOperationException($"筛选发送范围错误：{ViewModel.Scope} / {ViewModel.SendQueue().Count} / {ViewModel.VisibleJobs.Count}。");
        checks["send_scope_selection"] = true;
        SearchInput.Text = ""; ScopeInput.SelectedIndex = 1;

        var fixture = Path.Combine(output, "feature-fixtures"); Directory.CreateDirectory(fixture);
        var pdf = Path.Combine(fixture, "001_sample.pdf"); WriteSmokePdf(pdf);
        File.Copy(pdf, Path.Combine(fixture, "001_second.pdf"), true);
        var excel = Path.Combine(fixture, "students.xlsx");
        using (var book = new XLWorkbook())
        {
            var sheet = book.AddWorksheet("测试班级");
            sheet.Cell(1, 1).Value = "邮箱"; sheet.Cell(1, 2).Value = "学号"; sheet.Cell(1, 3).Value = "姓名";
            sheet.Cell(2, 1).Value = "sample@example.com"; sheet.Cell(2, 2).Value = "001"; sheet.Cell(2, 3).Value = "测试学生";
            book.SaveAs(excel);
        }
        var imported = ExcelImport.Inspect(excel).Single();
        if (imported.EmailColumn != 1 || imported.IdColumn != 2 || imported.NameColumn != 3) throw new InvalidOperationException("表头识别失败。");
        var store = new WorkspaceStore(Path.Combine(fixture, "library"));
        var template = new SavedTemplate("smoke", "示例模板", "{姓名}的资料", "{学号}同学：\n\n请查收两份 PDF 附件。");
        store.SaveTemplate(template);
        var value = new MailSettings { AttachmentFolder = fixture, StudentInfoExcel = excel, ExcelWorksheet = imported.Name,
            ExcelStudentIdCol = imported.IdColumn, ExcelNameCol = imported.NameColumn, ExcelEmailCol = imported.EmailColumn,
            EmailTemplate = store.Materialize(template), SenderEmail = "sender@example.com", GroupAttachments = true,
            StateFile = Path.Combine(fixture, "state.json") };
        var jobs = BatchPlanner.Prepare(value, value.StateFile);
        if (jobs.Count != 1 || jobs[0].AttachmentCount != 2 || jobs[0].Name != "测试学生") throw new InvalidOperationException("导入与合并附件没有贯通。");
        store.SaveTask(new("smoke-task", "示例任务", DateTimeOffset.Now, value with { SenderPassword = "never-persist-this" }, 1, 0, 0));
        if (store.Tasks.Single().Settings.SenderPassword.Length != 0) throw new InvalidOperationException("任务保存包含授权码。");
        checks["excel_template_task_group_flow"] = true;
        settings = value; FillForm(value, value.StateFile);
        await Task.Delay(150); ViewModel.LoadDemo(jobs, value);
        await Task.Delay(200); await CaptureAsync(output, "winui3-grouped-attachments.png");

        smokeLibraryDirectory = Path.Combine(fixture, "library");
        smokeDialogAction = async dialog =>
        {
            var panel = (StackPanel)((ScrollViewer)dialog.Content).Content;
            if (((ComboBox)panel.Children[2]).SelectedIndex != 1 || ((ComboBox)panel.Children[4]).SelectedIndex != 0)
                throw new InvalidOperationException("导入向导没有正确设置列映射。");
            await Task.Delay(200); await CaptureAsync(output, "winui3-excel-import.png", dialog);
            return ContentDialogResult.Primary;
        };
        await ExcelWizard(excel, true);
        if (Collect().ExcelStudentIdCol != 2 || Collect().ExcelWorksheet != imported.Name) throw new InvalidOperationException("导入映射没有应用到设置。");
        checks["excel_wizard_ui"] = true;
        smokeDialogAction = async dialog =>
        {
            var panel = (StackPanel)((ScrollViewer)dialog.Content).Content;
            ((TextBox)panel.Children[2]).Text = "界面编辑的模板";
            ((TextBox)panel.Children[3]).Text = "{姓名}的通知";
            ((TextBox)panel.Children[5]).Text = "学号 {学号}，请查收。";
            await Task.Delay(150);
            var preview = (TextBlock)((ScrollViewer)panel.Children[6]).Content;
            if (!preview.Text.Contains("测试学生的通知") || !preview.Text.Contains("学号 001，请查收。")) throw new InvalidOperationException("模板实时替换没有生效：" + preview.Text);
            await CaptureAsync(output, "winui3-template-editor.png", dialog);
            return ContentDialogResult.Primary;
        };
        await EditTemplatesAsync(true);
        if (!File.ReadAllText(TemplateInput.Text).Contains("{学号}")) throw new InvalidOperationException("模板没有保存并应用。");
        checks["template_editor_ui"] = true;
        smokeDialogAction = async dialog =>
        {
            var list = (ListView)((StackPanel)dialog.Content).Children[0]; list.SelectedIndex = 0;
            await Task.Delay(150); await CaptureAsync(output, "winui3-task-history.png", dialog);
            return ContentDialogResult.Secondary;
        };
        await ManageTasksAsync(true);
        if (taskId == "smoke-task" || Collect().StateFile != value.StateFile) throw new InvalidOperationException("任务复制没有保留发送日志。");
        checks["task_copy_preserves_journal"] = true;
        smokeLibraryDirectory = null;

        smokeDialogAction = async dialog =>
        {
            var panel = (StackPanel)dialog.Content;
            var selector = (ComboBox)panel.Children[0];
            var tools = (StackPanel)panel.Children[1];
            var page = (NumberBox)tools.Children[1]; var zoom = (ComboBox)tools.Children[3];
            var note = (TextBlock)panel.Children[2]; var scroll = (ScrollViewer)panel.Children[3]; var image = (Image)scroll.Content;
            async Task WaitFor(Func<bool> condition)
            {
                for (var i = 0; i < 50 && !condition(); i++) await Task.Delay(100);
                if (!condition()) throw new InvalidOperationException("PDF 原生预览没有完成：" + note.Text);
            }
            await WaitFor(() => image.Source is not null);
            if (page.Maximum != 2) throw new InvalidOperationException("PDF 页数错误。");
            page.Value = 2;
            await WaitFor(() => note.Text.StartsWith("第 2 /"));
            zoom.SelectedIndex = 3;
            await WaitFor(() => image.Width > 450);
            await CaptureAsync(output, "winui3-pdf-preview.png", dialog);
            selector.SelectedIndex = 1;
            await WaitFor(() => page.Value == 1 && image.Source is not null);
            checks["pdf_page_zoom_attachment_switch"] = true;
            return ContentDialogResult.None;
        };
        try { await PdfPreviewAsync(jobs[0]); }
        finally { smokeDialogAction = null; }
    }

    private static void WriteSmokePdf(string path)
    {
        var objects = new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << /Font << /F1 5 0 R >> >> /Contents 7 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "", "" };
        for (var i = 5; i < 7; i++)
        {
            var content = $"BT /F1 16 Tf 30 340 Td (PDF preview page {i - 4}) Tj ET\n";
            objects[i] = $"<< /Length {content.Length} >>\nstream\n{content}endstream";
        }
        var builder = new StringBuilder("%PDF-1.4\n"); var offsets = new List<int>();
        for (var i = 0; i < objects.Length; i++)
        { offsets.Add(builder.Length); builder.Append($"{i + 1} 0 obj\n{objects[i]}\nendobj\n"); }
        var xref = builder.Length;
        builder.Append($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets) builder.Append(offset.ToString("D10") + " 00000 n \n");
        builder.Append($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        File.WriteAllText(path, builder.ToString(), Encoding.ASCII);
    }
}
