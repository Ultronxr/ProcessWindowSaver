using System.Text;
using ProcessWindowSaver.Model;

namespace ProcessWindowSaver.Util;

public static class FileUtils {
    private const string PROCESS_INFO_FILE_PATH = "processInfo_{datetime}.txt";
    private const string FIELDS_SPLITTER = ":::";

    // 保存到txt文件
    public static void SaveToFile(List<ProcessInfo> processInfos) {
        if (processInfos.Count == 0) {
            Console.WriteLine("没有符合条件的进程信息，无法保存到txt文件！");
            return;
        }
        string filepath = PROCESS_INFO_FILE_PATH.Replace("{datetime}", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        using (StreamWriter writer = new StreamWriter(filepath, false, Encoding.UTF8)) {
            foreach (var info in processInfos) {
                writer.WriteLine(
                    $"{info.ProcessName}{FIELDS_SPLITTER}" +
                    $"{info.FilePath}{FIELDS_SPLITTER}" +
                    $"{info.CommandLine}{FIELDS_SPLITTER}" +
                    $"{info.X}{FIELDS_SPLITTER}" +
                    $"{info.Y}{FIELDS_SPLITTER}" +
                    $"{info.Width}{FIELDS_SPLITTER}" +
                    $"{info.Height}{FIELDS_SPLITTER}" +
                    $"{info.WindowTitle}");
            }
        }

        Console.WriteLine($"信息已保存到txt文件: {Path.GetFullPath(filepath)}");
    }

    // 列出所有的进程信息文件路径
    public static List<string> ListProcessInfoFiles() {
        string pattern = PROCESS_INFO_FILE_PATH.Replace("_{datetime}", "*");
        string currentDirectory = Directory.GetCurrentDirectory();
        string[] files = Directory.GetFiles(currentDirectory, pattern);
        return files.OrderByDescending(f => File.GetLastWriteTime(f)).ToList();
    }

    // 读取文件
    public static List<ProcessInfo> ReadFromFile(string filepath) {
        if (!File.Exists(filepath)) {
            throw new FileNotFoundException("文件不存在！");
        }
        if (!filepath.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("文件格式不正确，仅支持txt文件！");
        }

        List<ProcessInfo> processInfoList = new List<ProcessInfo>();
        using (StreamReader reader = new StreamReader(filepath, Encoding.UTF8)) {
            while (reader.Peek() != -1) {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) {
                    continue;
                }

                string[] fields = line.Split(FIELDS_SPLITTER);
                if (fields.Length != 8 || (fields[0] == "进程名称" && fields[1] == "文件路径")) {
                    continue;
                }

                ProcessInfo info = new ProcessInfo() {
                    ProcessName = fields[0],
                    FilePath = fields[1],
                    CommandLine = fields[2],
                    X = int.Parse(fields[3]),
                    Y = int.Parse(fields[4]),
                    Width = int.Parse(fields[5]),
                    Height = int.Parse(fields[6]),
                    WindowTitle = fields[7],
                };
                processInfoList.Add(info);
            }
        }

        return processInfoList;
    }
}