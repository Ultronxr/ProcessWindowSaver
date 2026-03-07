using System.Text;

namespace ProcessWindowSaver.Util;

/// <summary>
/// 命令行字符串解析与格式化辅助工具。
/// 主要用于：
/// 1. 从完整命令行中拆出程序路径与参数；
/// 2. 在恢复阶段把参数重新打印成便于观察的文本。
/// </summary>
public static class CommandLineUtils {
    /// <summary>
    /// 将完整命令行拆分为参数列表。
    /// 这是一个轻量解析器，支持最常见的双引号包裹参数场景。
    /// </summary>
    public static List<string> ParseArguments(string? commandLine) {
        List<string> arguments = new();
        if (string.IsNullOrWhiteSpace(commandLine)) {
            return arguments;
        }

        StringBuilder current = new();
        bool inQuotes = false;

        foreach (char ch in commandLine) {
            if (ch == '"') {
                inQuotes = !inQuotes;
                continue;
            }

            if (char.IsWhiteSpace(ch) && !inQuotes) {
                if (current.Length > 0) {
                    arguments.Add(current.ToString());
                    current.Clear();
                }

                continue;
            }

            current.Append(ch);
        }

        if (current.Length > 0) {
            arguments.Add(current.ToString());
        }

        return arguments;
    }

    /// <summary>
    /// 将参数列表重新格式化为可阅读的命令行文本。
    /// 仅用于日志输出，不参与真正的进程启动。
    /// </summary>
    public static string FormatArguments(IEnumerable<string> arguments) {
        return string.Join(" ", arguments.Select(QuoteIfNeeded));
    }

    /// <summary>
    /// 如果参数中包含空白字符，则自动补上双引号，避免日志输出产生歧义。
    /// </summary>
    private static string QuoteIfNeeded(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return "\"\"";
        }

        return value.Any(char.IsWhiteSpace)
            ? $"\"{value.Replace("\"", "\\\"")}\""
            : value;
    }
}
