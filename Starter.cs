using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using ProcessWindowSaver.Model;
using ProcessWindowSaver.Util;

namespace ProcessWindowSaver;

/// <summary>
/// 根据保存的进程快照重新启动应用，并恢复窗口位置与大小。
/// </summary>
[SupportedOSPlatform("windows")]
public class Starter {
    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    /// <summary>
    /// 不改变窗口 Z 轴顺序。
    /// </summary>
    private const uint SWP_NOZORDER = 0x0004;

    /// <summary>
    /// 在设置位置时确保窗口可见。
    /// </summary>
    private const uint SWP_SHOWWINDOW = 0x0040;

    /// <summary>
    /// 最大化窗口。
    /// </summary>
    private const int SW_MAXIMIZE = 3;

    /// <summary>
    /// 还原窗口。
    /// 这里配合最大化再还原，是为了尽量避免目标窗口处于最小化状态。
    /// </summary>
    private const int SW_RESTORE = 9;

    /// <summary>
    /// 恢复入口。
    /// 会在当前目录中查找快照 JSON 文件，并允许用户选择要恢复的文件。
    /// </summary>
    public static void Start() {
        Console.WriteLine("==========应用程序启动器==========");

        List<string> filepathList = FileUtils.ListProcessInfoFiles();
        if (filepathList.Count == 0) {
            Console.WriteLine("当前目录下未找到 processInfo_*.json 文件。");
            Console.WriteLine("请输入文件路径（可直接拖入窗口），或回车退出：");
            string? line = Console.ReadLine()?.Trim();
            if (!string.IsNullOrWhiteSpace(line)) {
                ReadProcessInfoAndStart(line);
            }
        } else if (filepathList.Count == 1) {
            ReadProcessInfoAndStart(filepathList[0]);
        } else {
            Console.WriteLine("当前目录下发现多个 processInfo_*.json 文件，请输入编号：");
            for (int i = 1; i <= filepathList.Count; i++) {
                Console.WriteLine($"[{i}] {filepathList[i - 1]}");
            }

            string? line = Console.ReadLine()?.Trim();
            if (int.TryParse(line, out int lineInt) && lineInt > 0 && lineInt <= filepathList.Count) {
                ReadProcessInfoAndStart(filepathList[lineInt - 1]);
            } else {
                Console.WriteLine("错误的文件编号！");
            }
        }

        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }

    /// <summary>
    /// 读取指定快照文件，并依次恢复其中记录的所有进程。
    /// </summary>
    private static void ReadProcessInfoAndStart(string filepath) {
        try {
            List<ProcessInfo> processInfos = FileUtils.ReadFromFile(filepath);
            foreach (ProcessInfo processInfo in processInfos) {
                StartProcess(processInfo);
            }
        } catch (Exception ex) {
            Console.WriteLine(ex.Message);
        }
    }

    /// <summary>
    /// 按单条进程快照恢复程序与窗口。
    /// </summary>
    private static void StartProcess(ProcessInfo processInfo) {
        if (string.IsNullOrWhiteSpace(processInfo.FilePath)) {
            return;
        }

        Console.WriteLine($"正在启动: {processInfo.ProcessName}");

        // 默认先使用主程序路径作为启动目标，后续再根据恢复模式覆盖参数。
        string appPath = processInfo.FilePath;
        List<string> launchArguments = new();

        // PotPlayer 特殊处理：如果已经识别到当前媒体文件，则优先直接打开该媒体文件，
        // 这样可以避免回到最初启动时的旧视频。
        if (processInfo.ResumeMode == ProcessInfo.PotPlayerCurrentMediaResumeMode &&
            !string.IsNullOrWhiteSpace(processInfo.CurrentMediaPath) &&
            File.Exists(processInfo.CurrentMediaPath)) {
            launchArguments.Add(processInfo.CurrentMediaPath);
            Console.WriteLine($"使用 PotPlayer 当前播放文件恢复: {processInfo.CurrentMediaPath}");
        } else {
            // 如果当前媒体文件不存在或不可用，则退回原始命令行恢复逻辑。
            if (processInfo.ResumeMode == ProcessInfo.PotPlayerCurrentMediaResumeMode &&
                !string.IsNullOrWhiteSpace(processInfo.CurrentMediaPath)) {
                Console.WriteLine("当前播放文件不可用，已回退到原始命令恢复。");
            }

            List<string> parsedCommandLine = CommandLineUtils.ParseArguments(processInfo.OriginalCommandLine);
            if (parsedCommandLine.Count > 0) {
                appPath = parsedCommandLine[0];
                launchArguments.AddRange(parsedCommandLine.Skip(1));
            }
        }

        // 如果命令行解析失败，则继续使用保存时记录的可执行文件路径。
        if (string.IsNullOrWhiteSpace(appPath)) {
            appPath = processInfo.FilePath;
        }

        Console.WriteLine($"执行文件: {appPath}");
        if (launchArguments.Count > 0) {
            Console.WriteLine($"参数: {CommandLineUtils.FormatArguments(launchArguments)}");
        }

        try {
            ProcessStartInfo startInfo = new() {
                FileName = appPath,
            };

            // 使用 ArgumentList 可以避免手工拼接参数字符串时的引号转义问题。
            foreach (string argument in launchArguments) {
                startInfo.ArgumentList.Add(argument);
            }

            Process process = new() {
                StartInfo = startInfo,
            };
            process.Start();

            Console.WriteLine("进程已启动，等待窗口初始化...");

            // 尝试等待目标进程进入空闲，以便主窗口句柄尽快可用。
            process.WaitForInputIdle(5000);

            // 某些 GUI 程序在进入输入空闲后，主窗口仍需要一点时间完成布局。
            Thread.Sleep(3000);
            process.Refresh();

            if (process.MainWindowHandle != IntPtr.Zero) {
                Console.WriteLine($"设置窗口位置: ({processInfo.X}, {processInfo.Y})，大小: {processInfo.Width}x{processInfo.Height}");

                // 先最大化再还原，尽量让窗口退出最小化/隐藏状态。
                ShowWindow(process.MainWindowHandle, SW_MAXIMIZE);
                ShowWindow(process.MainWindowHandle, SW_RESTORE);

                // 恢复窗口大小和位置。
                SetWindowPos(
                    process.MainWindowHandle,
                    IntPtr.Zero,
                    processInfo.X,
                    processInfo.Y,
                    processInfo.Width,
                    processInfo.Height,
                    SWP_NOZORDER | SWP_SHOWWINDOW);

                Console.WriteLine("窗口设置完成！");
            } else {
                Console.WriteLine("未获取到窗口句柄，应用可能没有图形界面。");
            }
        } catch (Exception ex) {
            Console.WriteLine($"错误: {ex.Message}");
        }
    }
}
