using System.Text.Json;
using ProcessWindowSaver.Model;

namespace ProcessWindowSaver.Util;

/// <summary>
/// 负责进程窗口快照的持久化读写。
/// 目前统一使用 JSON 文件，不再维护旧的 TXT 存储格式。
/// </summary>
public static class FileUtils {
    /// <summary>
    /// 保存文件名模板。
    /// 最终会生成形如 processInfo_20260307-013500.json 的文件。
    /// </summary>
    private const string ProcessInfoFilePath = "processInfo_{datetime}.json";

    /// <summary>
    /// JSON 序列化配置。
    /// 使用缩进格式，方便人工查看和调试。
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true,
    };

    /// <summary>
    /// 将进程快照集合保存为 JSON 文件。
    /// </summary>
    public static void SaveToFile(List<ProcessInfo> processInfos) {
        if (processInfos.Count == 0) {
            Console.WriteLine("没有符合条件的进程信息，未生成保存文件。");
            return;
        }

        string filepath = ProcessInfoFilePath.Replace("{datetime}", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        string json = JsonSerializer.Serialize(processInfos, JsonOptions);
        File.WriteAllText(filepath, json);

        Console.WriteLine($"信息已保存到: {Path.GetFullPath(filepath)}");
    }

    /// <summary>
    /// 枚举当前目录下所有快照文件，并按最后修改时间倒序返回。
    /// 这样恢复时默认更容易命中最新一次保存结果。
    /// </summary>
    public static List<string> ListProcessInfoFiles() {
        string pattern = ProcessInfoFilePath.Replace("_{datetime}", "*");
        string currentDirectory = Directory.GetCurrentDirectory();
        string[] files = Directory.GetFiles(currentDirectory, pattern);
        return files.OrderByDescending(File.GetLastWriteTime).ToList();
    }

    /// <summary>
    /// 从 JSON 文件读取进程快照集合。
    /// </summary>
    public static List<ProcessInfo> ReadFromFile(string filepath) {
        if (!File.Exists(filepath)) {
            throw new FileNotFoundException("文件不存在！");
        }

        if (!filepath.EndsWith(".json", StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("文件格式不正确，仅支持 json 文件！");
        }

        string json = File.ReadAllText(filepath);
        List<ProcessInfo>? processInfoList = JsonSerializer.Deserialize<List<ProcessInfo>>(json, JsonOptions);
        return processInfoList ?? new List<ProcessInfo>();
    }
}
