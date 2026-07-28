using System.Diagnostics;
using System.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace NarouDownloaderGui;

public sealed partial class MainPage : Page
{
    private static readonly Encoding DownloaderOutputEncoding = CreateDownloaderOutputEncoding();

    private Process? _activeProcess;
    private string? _lastOutputPath;
    private bool _stopRequested;

    public MainPage()
    {
        InitializeComponent();

        OutputFolderTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        FileNameTextBox.Text = "novel.txt";
    }

    private async void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (App.MainWindow is null)
        {
            ShowStatus(InfoBarSeverity.Error, "保存先を選択できません", "アプリのウィンドウを取得できませんでした。");
            return;
        }

        var picker = new FolderPicker
        {
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary
        };
        picker.FileTypeFilter.Add("*");
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(App.MainWindow));

        var folder = await picker.PickSingleFolderAsync();
        if (folder is not null)
        {
            OutputFolderTextBox.Text = folder.Path;
        }
    }

    private async void DownloadButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildRequest(out var request, out var validationMessage))
        {
            ShowStatus(InfoBarSeverity.Warning, "入力内容を確認してください", validationMessage);
            return;
        }

        if (File.Exists(request.OutputPath))
        {
            var overwriteDialog = new ContentDialog
            {
                XamlRoot = XamlRoot,
                Title = "ファイルを上書きしますか？",
                Content = $"{request.OutputPath}\n\n同名のファイルが既に存在します。",
                PrimaryButtonText = "上書き",
                CloseButtonText = "キャンセル",
                DefaultButton = ContentDialogButton.Close
            };

            if (await overwriteDialog.ShowAsync() != ContentDialogResult.Primary)
            {
                return;
            }
        }

        await RunDownloaderAsync(request);
    }

    private void StopButton_Click(object sender, RoutedEventArgs e)
    {
        var process = _activeProcess;
        if (process is null || process.HasExited)
        {
            return;
        }

        _stopRequested = true;
        StopButton.IsEnabled = false;
        ShowStatus(InfoBarSeverity.Warning, "停止しています", "ダウンローダーの終了を待っています。");

        try
        {
            process.Kill(entireProcessTree: true);
        }
        catch (InvalidOperationException)
        {
            // The process exited between the state check and the kill request.
        }
        catch (Exception ex)
        {
            AppendLog($"\n停止処理に失敗しました: {ex.Message}\n");
        }
    }

    private void OpenOutputButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastOutputPath) || !File.Exists(_lastOutputPath))
        {
            ShowStatus(InfoBarSeverity.Warning, "ファイルが見つかりません", "保存されたファイルを確認できませんでした。");
            OpenOutputButton.IsEnabled = false;
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/select,\"{_lastOutputPath}\"",
            UseShellExecute = true
        });
    }

    private void ClearLogButton_Click(object sender, RoutedEventArgs e)
    {
        LogTextBox.Text = string.Empty;
    }

    private void NovelUrlTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (StatusInfoBar.Severity == InfoBarSeverity.Warning)
        {
            ShowStatus(InfoBarSeverity.Informational, "準備完了", "作品URLと保存先を指定してください。");
        }
    }

    private void Page_Unloaded(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_activeProcess is { HasExited: false })
            {
                _activeProcess.Kill(entireProcessTree: true);
            }
        }
        catch
        {
            // The application is closing; there is no useful recovery action.
        }
    }

    private bool TryBuildRequest(out DownloadRequest request, out string validationMessage)
    {
        request = default;
        validationMessage = string.Empty;

        var urlText = NovelUrlTextBox.Text.Trim();
        if (!Uri.TryCreate(urlText, UriKind.Absolute, out var uri) ||
            uri.Scheme != Uri.UriSchemeHttps ||
            (uri.Host is not "ncode.syosetu.com" and not "novel18.syosetu.com") ||
            !uri.AbsolutePath.StartsWith("/n", StringComparison.OrdinalIgnoreCase))
        {
            validationMessage = "対応サイトの作品トップページURLを入力してください。";
            return false;
        }

        var outputFolder = OutputFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
        {
            validationMessage = "存在する保存先フォルダーを選択してください。";
            return false;
        }

        var fileName = FileNameTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(fileName))
        {
            validationMessage = "保存ファイル名を入力してください。";
            return false;
        }

        if (Path.GetFileName(fileName) != fileName ||
            fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            validationMessage = "保存ファイル名にはフォルダー名や使用できない記号を含めないでください。";
            return false;
        }

        if (!fileName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
        {
            fileName += ".txt";
            FileNameTextBox.Text = fileName;
        }

        var startPageValue = StartPageNumberBox.Value;
        var startPage = double.IsNaN(startPageValue) ? 1 : (int)startPageValue;
        if (startPage < 1)
        {
            validationMessage = "開始話には1以上の数値を指定してください。";
            return false;
        }

        request = new DownloadRequest(uri.AbsoluteUri, Path.Combine(outputFolder, fileName), startPage);
        return true;
    }

    private async Task RunDownloaderAsync(DownloadRequest request)
    {
        var executablePath = Path.Combine(AppContext.BaseDirectory, "Tools", "na6dl.exe");
        if (!File.Exists(executablePath))
        {
            ShowStatus(InfoBarSeverity.Error, "ダウンローダーが見つかりません", executablePath);
            return;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = DownloaderOutputEncoding,
            StandardErrorEncoding = DownloaderOutputEncoding
        };

        if (request.StartPage > 1)
        {
            startInfo.ArgumentList.Add($"-s{request.StartPage}");
        }

        startInfo.ArgumentList.Add(request.Url);
        startInfo.ArgumentList.Add(request.OutputPath);

        using var process = new Process { StartInfo = startInfo };
        _activeProcess = process;
        _lastOutputPath = null;
        _stopRequested = false;
        SetRunningState(true);
        LogTextBox.Text = string.Empty;
        AppendLog($"保存先: {request.OutputPath}\n開始話: {request.StartPage}\n\n");
        ShowStatus(InfoBarSeverity.Informational, "ダウンロード中", "完了までアプリを閉じずにお待ちください。");

        try
        {
            if (!process.Start())
            {
                throw new InvalidOperationException("ダウンローダーを開始できませんでした。");
            }

            var outputTask = PumpOutputAsync(process.StandardOutput);
            var errorTask = PumpOutputAsync(process.StandardError);

            await process.WaitForExitAsync();
            await Task.WhenAll(outputTask, errorTask);

            if (_stopRequested)
            {
                ShowStatus(InfoBarSeverity.Warning, "停止しました", "ダウンロードはユーザー操作で停止されました。");
            }
            else if (process.ExitCode == 0 && File.Exists(request.OutputPath))
            {
                _lastOutputPath = request.OutputPath;
                OpenOutputButton.IsEnabled = true;
                ShowStatus(InfoBarSeverity.Success, "保存しました", request.OutputPath);
            }
            else
            {
                ShowStatus(
                    InfoBarSeverity.Error,
                    "ダウンロードに失敗しました",
                    $"終了コード: {process.ExitCode}。実行ログを確認してください。");
            }
        }
        catch (Exception ex)
        {
            AppendLog($"\nエラー: {ex.Message}\n");
            ShowStatus(InfoBarSeverity.Error, "実行エラー", ex.Message);
        }
        finally
        {
            _activeProcess = null;
            SetRunningState(false);
        }
    }

    private async Task PumpOutputAsync(StreamReader reader)
    {
        var buffer = new char[256];

        while (true)
        {
            var count = await reader.ReadAsync(buffer);
            if (count == 0)
            {
                break;
            }

            var text = new string(buffer, 0, count)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Replace('\r', '\n');

            DispatcherQueue.TryEnqueue(() => AppendLog(text));
        }
    }

    private void AppendLog(string text)
    {
        LogTextBox.Text += text;
        LogTextBox.Select(LogTextBox.Text.Length, 0);
    }

    private void SetRunningState(bool isRunning)
    {
        DownloadButton.IsEnabled = !isRunning;
        StopButton.IsEnabled = isRunning;
        NovelUrlTextBox.IsEnabled = !isRunning;
        OutputFolderTextBox.IsEnabled = !isRunning;
        FileNameTextBox.IsEnabled = !isRunning;
        StartPageNumberBox.IsEnabled = !isRunning;
        DownloadProgressBar.Visibility = isRunning ? Visibility.Visible : Visibility.Collapsed;
    }

    private void ShowStatus(InfoBarSeverity severity, string title, string message)
    {
        StatusInfoBar.Severity = severity;
        StatusInfoBar.Title = title;
        StatusInfoBar.Message = message;
        StatusInfoBar.IsOpen = true;
    }

    private static Encoding CreateDownloaderOutputEncoding()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(932);
    }

    private readonly record struct DownloadRequest(string Url, string OutputPath, int StartPage);
}
