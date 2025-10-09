using System.Diagnostics;
using System.Runtime.InteropServices;
using ProcessWindowSaver.Util;
using ProcessWindowSaver.Model;

namespace ProcessWindowSaver;

public class Starter {
    // 引入Windows API函数
    [DllImport("user32.dll")]
    static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    // 常量定义
    const uint SWP_NOSIZE = 0x0001;
    const uint SWP_NOZORDER = 0x0004;
    const uint SWP_SHOWWINDOW = 0x0040;
    const int SW_MAXIMIZE = 3;
    const int SW_RESTORE = 9;

    public static void Start() {
        Console.WriteLine("==========应用程序启动器==========");
        
        List<string> filepathList = FileUtils.ListProcessInfoFiles();
        if (filepathList.Count == 0) {
            Console.WriteLine("同路径下未找到任何 processInfo.txt 文件！");
        } else if (filepathList.Count == 1) {
            List<ProcessInfo> processInfos = FileUtils.ReadFromFile(filepathList[0]);
            foreach (ProcessInfo processInfo in processInfos) {
                StartProcess(processInfo);
            }
        } else {
            Console.WriteLine("同路径下发现多个 processInfo.txt 文件，请指定一个文件编号：");
            for(int i = 1; i <= filepathList.Count; i++) {
                Console.WriteLine($"[{i}] {filepathList[i - 1]}");
            }
            string? line = Console.ReadLine()?.Trim();
            if (int.TryParse(line, out int lineInt) && lineInt > 0 && lineInt <= filepathList.Count) {
                List<ProcessInfo> processInfos = FileUtils.ReadFromFile(filepathList[lineInt]);
                foreach (ProcessInfo processInfo in processInfos) {
                    StartProcess(processInfo);
                }
            } else {
                Console.WriteLine("错误的文件编号！");
            }
        }
        
        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }

    private static void StartProcess(ProcessInfo processInfo) {
        if (string.IsNullOrWhiteSpace(processInfo.FilePath) || processInfo.FilePath == "无法获取") {
            return;
        }
        
        Console.WriteLine($"正在启动: {processInfo.ProcessName}");
        // 设置要启动的应用程序路径和参数
        string appPath;
        string appArgs;
        if (string.IsNullOrWhiteSpace(processInfo.CommandLine) || processInfo.CommandLine == "无参数") {
            // 如果命令行是空的，那么只取应用程序路径即可
            appPath = processInfo.FilePath;
            appArgs = "";
        } else {
            // 需要从命令行中解析出自定义参数
            if (processInfo.CommandLine.StartsWith("\"")) {
                // 以双引号开头，那么第一对双引号内的一定是程序路径，后面的都是参数
                int secondQuoteIndex = processInfo.CommandLine.IndexOf("\"", 1, StringComparison.Ordinal);
                appPath = processInfo.CommandLine.Substring(0, secondQuoteIndex + 1);
                appArgs = processInfo.CommandLine.Substring(secondQuoteIndex + 1);
            } else {
                // 不以双引号开头，说明程序的文件路径里没有空格（有空格一定会用双引号括起来），那么第一个空格前的一定是程序路径，后面的都是参数
                string[] cmdArr = processInfo.CommandLine.Split(" ");
                appPath = cmdArr[0];
                appArgs = cmdArr.Length > 1 ? " " + string.Join(" ", cmdArr[1..]) : "";
            }
        }
        
        int x = processInfo.X; // 窗口左上角X坐标
        int y = processInfo.Y; // 窗口左上角Y坐标
        int width = processInfo.Width; // 窗口宽度
        int height = processInfo.Height; // 窗口高度
        
        Console.WriteLine($"正在启动: {appPath}");
        if (!string.IsNullOrEmpty(appArgs)) {
            Console.WriteLine($"参数: {appArgs}");
        }
        
        try {
            // 启动进程
            Process process = new Process();
            process.StartInfo.FileName = appPath;
            process.StartInfo.Arguments = appArgs;
            process.Start();
        
            Console.WriteLine("进程已启动，等待窗口初始化...");
        
            // 等待进程创建窗口
            process.WaitForInputIdle(5000);
        
            // 给窗口一些时间来完全初始化
            Thread.Sleep(3000);
        
            // 设置窗口位置和大小
            if (process.MainWindowHandle != IntPtr.Zero) {
                Console.WriteLine($"设置窗口位置: ({x}, {y})，大小: {width}x{height}");
        
                // 先最大化然后还原，确保窗口不是最小化状态
                ShowWindow(process.MainWindowHandle, SW_MAXIMIZE);
                ShowWindow(process.MainWindowHandle, SW_RESTORE);
        
                // 设置窗口位置和大小
                SetWindowPos(process.MainWindowHandle, IntPtr.Zero, x, y, width, height,
                    SWP_NOZORDER | SWP_SHOWWINDOW);
        
                Console.WriteLine("窗口设置完成!");
            } else {
                Console.WriteLine("无法获取窗口句柄，应用程序可能没有图形界面。");
            }
        } catch (Exception ex) {
            Console.WriteLine($"错误: {ex.Message}");
        }
    }
}