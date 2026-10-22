using Microsoft.UI.Xaml;

namespace PdfMailer.App;

public partial class App : Application
{
    private Window? window;
    public App()
    {
        InitializeComponent();
        UnhandledException += (_, args) =>
        {
            // 无控制台的桌面程序也保留可诊断的启动错误。
            try
            {
                var command = Environment.GetCommandLineArgs();
                var index = Array.IndexOf(command, "--smoke-test");
                var smoke = index >= 0 && index + 1 < command.Length ? command[index + 1] : null;
                var path = smoke ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PdfMailer");
                Directory.CreateDirectory(path);
                var diagnostic = args.Message + "\n" + args.Exception;
                File.WriteAllText(Path.Combine(path, "last-error.txt"), diagnostic);
                if (smoke is not null)
                    File.WriteAllText(Path.Combine(path, "smoke-report.json"), System.Text.Json.JsonSerializer.Serialize(new { success = false, error = diagnostic }));
            }
            catch { }
        };
    }
    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        window = new MainWindow();
        window.Activate();
    }
}
