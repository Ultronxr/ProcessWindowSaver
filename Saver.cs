using ProcessWindowSaver.Util;
using ProcessWindowSaver.Model;

namespace ProcessWindowSaver;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

public class Saver {
    // 引入Windows API函数
    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll")]
    static extern bool QueryFullProcessImageName(IntPtr hProcess, int dwFlags,
        StringBuilder lpExeName, out int size);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;

        public int Width => Right - Left;
        public int Height => Bottom - Top;
    }

    public static void Save(string? processNameFilter = null) {
        Console.WriteLine("应用程序信息导出工具");
        Console.WriteLine("=====================");
        string[]? processNameFilterArr = string.IsNullOrWhiteSpace(processNameFilter)
            ? null
            : processNameFilter.Trim().ToLower().Split(',');

        // 获取所有可见窗口的进程信息
        List<ProcessInfo> processInfos = GetVisibleProcessesInfo(processNameFilterArr);

        // 保存到CSV文件
        FileUtils.SaveToFile(processInfos);

        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }

    // 获取所有可见窗口的进程信息
    static List<ProcessInfo> GetVisibleProcessesInfo(string[]? processNameFilterArr) {
        var processInfos = new List<ProcessInfo>();

        // 遍历所有进程
        foreach (Process process in Process.GetProcesses()) {
            try {
                // 检查进程是否有主窗口且可见
                if (process.MainWindowHandle != IntPtr.Zero && IsWindowVisible(process.MainWindowHandle)) {
                    string processName = process.ProcessName.Trim().ToLower();
                    Console.WriteLine(processName);
                    // processNameFilterArr != null 时表示启用了 filter ，此时需要判断进程名称是否在过滤白名单中
                    if (processNameFilterArr != null && !processNameFilterArr.Any(p => processName.Contains(p))) {
                        continue;
                    }

                    // 获取窗口位置和大小
                    if (GetWindowRect(process.MainWindowHandle, out RECT rect)) {
                        // 获取进程文件路径
                        string filePath = GetProcessFilePath(process);

                        // 获取命令行参数
                        string commandLine = GetCommandLineArgs(process);

                        // 创建进程信息对象
                        ProcessInfo info = new ProcessInfo {
                            ProcessName = process.ProcessName,
                            FilePath = filePath,
                            CommandLine = commandLine,
                            X = rect.Left,
                            Y = rect.Top,
                            Width = rect.Width,
                            Height = rect.Height,
                            WindowTitle = process.MainWindowTitle
                        };

                        processInfos.Add(info);
                        Console.WriteLine($"已获取: {process.ProcessName}");
                    }
                }
            } catch (Exception ex) {
                // 忽略无法访问的进程
                Console.WriteLine($"无法访问进程 {process.ProcessName}: {ex.Message}");
            }
        }

        return processInfos;
    }

    // 获取进程文件路径
    static string GetProcessFilePath(Process process) {
        try {
            // 尝试直接获取主模块文件名
            return process.MainModule?.FileName ?? "无法获取";
        } catch {
            try {
                // 如果上面方法失败，使用API获取
                StringBuilder buffer = new StringBuilder(1024);
                int size = buffer.Capacity;
                if (QueryFullProcessImageName(process.Handle, 0, buffer, out size)) {
                    return buffer.ToString();
                }
            } catch {
                // 如果所有方法都失败
                return "无法获取";
            }
        }

        return "无法获取";
    }

    // 获取命令行参数（需要引用System.Management）
    static string GetCommandLineArgs(Process process) {
        try {
            // 使用ManagementObjectSearcher获取命令行参数
            // 注意：需要添加对System.Management的引用
            using (var searcher = new System.Management.ManagementObjectSearcher(
                       $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}")) {
                foreach (System.Management.ManagementObject obj in searcher.Get()) {
                    return obj["CommandLine"]?.ToString() ?? "无参数";
                }
            }

            return "无参数";
        } catch {
            return "无法获取";
        }
    }
}