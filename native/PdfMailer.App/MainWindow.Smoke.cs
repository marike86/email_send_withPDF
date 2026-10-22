using System.Runtime.InteropServices.WindowsRuntime;
using System.Text.Json;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media.Imaging;
using PdfMailer.Core;
using Windows.Graphics.Imaging;
using Windows.Storage.Streams;

namespace PdfMailer.App;

public partial class MainWindow
{
    // 仅 --smoke-test 显式启用。使用演示数据，从不连接 SMTP 或改写用户配置。
    private async Task RunSmokeAsync(string output)
    {
        Directory.CreateDirectory(output);
        var checks = new Dictionary<string, object>();
        try
        {
            Root.RequestedTheme = ElementTheme.Light;
            await Task.Delay(700);
            Root.UpdateLayout();
            checks["root_width"] = Root.ActualWidth;
            checks["root_height"] = Root.ActualHeight;
            checks["preview_viewport"] = PreviewScroll.ViewportHeight;
            if (PreviewScroll.ViewportHeight <= 80) throw new InvalidOperationException("邮件预览没有获得有效的阅读高度。");
            await CaptureAsync(output, "winui3-workspace.png");
            StateFilter.SelectedIndex = 3;
            if (ViewModel.VisibleJobs.Count != 2) throw new InvalidOperationException("原生下拉筛选未生效。");
            StateFilter.SelectedIndex = 0;
            checks["native_combo_filter"] = true;
            var first = ViewModel.Jobs[0];
            ViewModel.LoadDemo([new MailJob
            {
                Filename = first.Filename, Name = first.Name, Email = first.Email, Subject = first.Subject,
                Body = string.Join("\n\n", Enumerable.Range(1, 50).Select(i => $"第 {i} 段：这是一封较长的邮件。阅读区域自动换行，并使用 WinUI 原生滚动条。")), Key = "long-preview"
            }]);
            await Task.Delay(250);
            Root.UpdateLayout();
            if (PreviewScroll.ScrollableHeight <= 0) throw new InvalidOperationException("长邮件没有形成原生滚动范围。");
            PreviewScroll.ChangeView(null, 220, null, true);
            await Task.Delay(200);
            checks["long_mail_scrollable_height"] = PreviewScroll.ScrollableHeight;
            checks["long_mail_vertical_offset"] = PreviewScroll.VerticalOffset;
            if (PreviewScroll.VerticalOffset <= 0) throw new InvalidOperationException("原生阅读滚动未生效。");
            await CaptureAsync(output, "winui3-long-preview.png");
            LoadDemo();
            Root.RequestedTheme = ElementTheme.Dark;
            await Task.Delay(250);
            await CaptureAsync(output, "winui3-dark.png");
            Root.RequestedTheme = ElementTheme.Light;
            Navigation.SelectedItem = Navigation.SettingsItem;
            await Task.Delay(250);
            await CaptureAsync(output, "winui3-settings.png");
            checks["success"] = true;
        }
        catch (Exception exception)
        { checks["success"] = false; checks["error"] = exception.ToString(); }
        finally
        {
            await File.WriteAllTextAsync(Path.Combine(output, "smoke-report.json"), JsonSerializer.Serialize(checks, Storage.Json));
            Close();
        }
    }

    private async Task CaptureAsync(string folder, string name)
    {
        var bitmap = new RenderTargetBitmap();
        await bitmap.RenderAsync(Root);
        var pixels = await bitmap.GetPixelsAsync();
        using var stream = new InMemoryRandomAccessStream();
        var encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, stream);
        encoder.SetPixelData(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied,
            (uint)bitmap.PixelWidth, (uint)bitmap.PixelHeight, 96, 96, pixels.ToArray());
        await encoder.FlushAsync();
        using var input = stream.AsStreamForRead();
        input.Position = 0;
        using var output = File.Create(Path.Combine(folder, name));
        await input.CopyToAsync(output);
    }
}
