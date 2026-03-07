using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.RegularExpressions;
using ProcessWindowSaver.Model;
using ProcessWindowSaver.Util;

namespace ProcessWindowSaver;

/// <summary>
/// 扫描当前可见窗口并生成可恢复快照。
/// 对普通程序保存原始命令行即可恢复；对 PotPlayer 会额外识别当前播放媒体路径。
/// </summary>
[SupportedOSPlatform("windows")]
public class Saver {
    /// <summary>
    /// 读取窗口矩形区域，用于获取窗口位置和大小。
    /// </summary>
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    /// <summary>
    /// 判断窗口当前是否可见。
    /// </summary>
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    /// <summary>
    /// 通过 Win32 API 查询进程可执行文件完整路径。
    /// </summary>
    [DllImport("kernel32.dll")]
    private static extern bool QueryFullProcessImageName(IntPtr hProcess, int dwFlags, StringBuilder lpExeName, out int size);

    /// <summary>
    /// Win32 窗口矩形结构，用于保存窗口左上角坐标以及宽高。
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;

        public int Width => Right - Left;
        public int Height => Bottom - Top;
    }

    /// <summary>
    /// 从 PotPlayer 的播放状态文件中提取出的简化会话信息。
    /// 其中最关键的是当前播放文件完整路径和状态文件的更新时间。
    /// </summary>
    private sealed class PotPlayerPlaylistState {
        public string PlaylistFilePath { get; init; } = string.Empty;
        public string CurrentMediaPath { get; init; } = string.Empty;
        public DateTime LastWriteTimeUtc { get; init; }
        public bool IsPreferredSessionFile { get; init; }
    }

    /// <summary>
    /// 识别为播放列表的扩展名集合，用于从命令行中展开播放列表内容。
    /// </summary>
    private static readonly HashSet<string> PlaylistExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".m3u", ".m3u8", ".pls", ".dpl", ".asx", ".wax", ".wvx", ".wpl", ".cue"
    };

    /// <summary>
    /// 识别为媒体文件的扩展名集合，用于过滤命令行候选和同目录扫描候选。
    /// </summary>
    private static readonly HashSet<string> MediaExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm", ".m4v", ".ts", ".m2ts", ".mpg", ".mpeg",
        ".mp3", ".flac", ".aac", ".wav", ".ogg", ".m4a", ".ape", ".wma"
    };

    /// <summary>
    /// PotPlayer 常见窗口标题后缀，用于剥离播放器名称只保留媒体标题。
    /// </summary>
    private static readonly string[] PotPlayerSuffixes = {
        " - PotPlayer",
        " - PotPlayerMini",
        " - PotPlayerMini64",
        " - PotPlayer64",
        " - Daum PotPlayer"
    };

    /// <summary>
    /// 去掉标题前部的短元信息，例如 [暂停]、(预览) 等内容。
    /// </summary>
    private static readonly Regex LeadingBracketMetadataRegex = new(@"^\s*[\[(][^\])]{1,40}[\])]\s*", RegexOptions.Compiled);
    /// <summary>
    /// 去掉标题尾部的短元信息。
    /// </summary>
    private static readonly Regex TrailingBracketMetadataRegex = new(@"\s*[\[(][^\])]{1,40}[\])]\s*$", RegexOptions.Compiled);
    /// <summary>
    /// 去掉标题中常见的分辨率、编码和帧率等噪声。
    /// </summary>
    private static readonly Regex ResolutionNoiseRegex = new(@"\s+(2160p|1080p|720p|540p|480p|x265|x264|hevc|av1|hdr10?\+?|\d{2,3}fps)\b.*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    /// <summary>
    /// 去掉标题中可能出现的播放时间文本。
    /// </summary>
    private static readonly Regex TimeNoiseRegex = new(@"\s+\d{1,2}:\d{2}(?::\d{2})?(\s*/\s*\d{1,2}:\d{2}(?::\d{2})?)?\s*$", RegexOptions.Compiled);

    /// <summary>
    /// 保存入口。
    /// 可选地按进程名白名单过滤需要导出的窗口。
    /// </summary>
    public static void Save(string? processNameFilter = null) {
        Console.WriteLine("==========应用程序信息导出工具==========");
        string[]? processNameFilterArr = string.IsNullOrWhiteSpace(processNameFilter)
            ? null
            : processNameFilter
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(item => item.ToLowerInvariant())
                .ToArray();

        List<ProcessInfo> processInfos = GetVisibleProcessesInfo(processNameFilterArr);
        FileUtils.SaveToFile(processInfos);

        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }

    /// <summary>
    /// 枚举当前所有可见主窗口，并提取可恢复的进程信息。
    /// </summary>
    private static List<ProcessInfo> GetVisibleProcessesInfo(string[]? processNameFilterArr) {
        List<ProcessInfo> processInfos = new();

        foreach (Process process in Process.GetProcesses()) {
            try {
                if (process.MainWindowHandle == IntPtr.Zero || !IsWindowVisible(process.MainWindowHandle)) {
                    continue;
                }

                string processName = process.ProcessName.Trim().ToLowerInvariant();
                Console.WriteLine($"扫描 {processName}");

                if (processNameFilterArr != null && !processNameFilterArr.Any(filter => processName.Contains(filter))) {
                    continue;
                }

                if (!GetWindowRect(process.MainWindowHandle, out RECT rect)) {
                    continue;
                }

                ProcessInfo info = new() {
                    ProcessName = process.ProcessName,
                    FilePath = GetProcessFilePath(process),
                    OriginalCommandLine = GetCommandLineArgs(process),
                    X = rect.Left,
                    Y = rect.Top,
                    Width = rect.Width,
                    Height = rect.Height,
                    WindowTitle = process.MainWindowTitle,
                };

                EnrichResumeState(info);
                processInfos.Add(info);
                Console.WriteLine($"已保存 {process.ProcessName}");
            } catch (Exception ex) {
                Console.WriteLine($"无法访问进程 {process.ProcessName}: {ex.Message}");
            }
        }

        return processInfos;
    }

    /// <summary>
    /// 为特定应用补充恢复所需的额外状态。
    /// 目前主要对 PotPlayer 增强当前播放媒体识别。
    /// </summary>
    private static void EnrichResumeState(ProcessInfo info) {
        if (!IsPotPlayerProcess(info.ProcessName)) {
            return;
        }

        if (TryResolvePotPlayerCurrentMedia(info, out string? currentMediaPath, out string details)) {
            info.CurrentMediaPath = currentMediaPath;
            info.ResumeMode = ProcessInfo.PotPlayerCurrentMediaResumeMode;
            Console.WriteLine($"PotPlayer 当前播放文件已识别: {currentMediaPath}");
            Console.WriteLine($"PotPlayer 识别详情: {details}");
            return;
        }

        Console.WriteLine("PotPlayer 当前播放文件未识别，将退回原始命令恢复。");
        Console.WriteLine($"PotPlayer 识别详情: {details}");
    }

    /// <summary>
    /// 判断当前进程是否为 PotPlayer 变体。
    /// </summary>
    private static bool IsPotPlayerProcess(string processName) {
        return processName.Contains("potplayer", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 识别 PotPlayer 当前播放媒体的完整路径。
    /// 识别顺序为：状态文件 -> 命令行/标题直接匹配 -> 同目录兜底匹配。
    /// </summary>
    private static bool TryResolvePotPlayerCurrentMedia(ProcessInfo info, out string? currentMediaPath, out string details) {
        currentMediaPath = null;

        List<string> titleCandidates = BuildPotPlayerTitleCandidates(info.WindowTitle);
        if (TryResolvePotPlayerCurrentMediaFromStateFiles(info.ProcessName, titleCandidates, out currentMediaPath, out string stateDetails)) {
            details = stateDetails;
            return true;
        }

        List<string> commandLineCandidates = CollectPotPlayerMediaCandidates(info.OriginalCommandLine);
        if (commandLineCandidates.Count == 0) {
            details = $"PotPlayer 状态文件未命中；原始命令行未解析出媒体候选。{stateDetails}";
            return false;
        }

        List<(string Path, int Score)> directMatches = ScoreCandidates(commandLineCandidates, titleCandidates);
        if (TryPickBestCandidate(directMatches, 70, out currentMediaPath, out string directMatchDetails)) {
            details = $"命令行候选匹配成功。{directMatchDetails} {stateDetails}".Trim();
            return true;
        }

        if (commandLineCandidates.Count == 1) {
            string primaryCandidate = commandLineCandidates[0];
            if (titleCandidates.Count == 0) {
                currentMediaPath = primaryCandidate;
                details = $"窗口标题为空，命令行仅包含一个媒体候选，直接采用。{stateDetails}".Trim();
                return true;
            }

            if (TryResolveFromSiblingFiles(primaryCandidate, titleCandidates, out currentMediaPath, out string siblingDetails)) {
                details = $"{siblingDetails} {stateDetails}".Trim();
                return true;
            }

            details = $"单候选与标题不一致，且同目录未找到更优匹配。{directMatchDetails} {stateDetails}".Trim();
            return false;
        }

        details = $"{directMatchDetails} {stateDetails}".Trim();
        return false;
    }

    /// <summary>
    /// 从 PotPlayer 的 *.dpl 状态文件中直接读取 playname 字段。
    /// 这是支持跨目录切换后仍能保存完整当前路径的关键实现。
    /// </summary>
    private static bool TryResolvePotPlayerCurrentMediaFromStateFiles(
        string processName,
        List<string> titleCandidates,
        out string? currentMediaPath,
        out string details) {
        currentMediaPath = null;

        List<PotPlayerPlaylistState> states = ReadPotPlayerPlaylistStates(processName).ToList();
        if (states.Count == 0) {
            details = "未找到 PotPlayer 播放状态文件。";
            return false;
        }

        List<(PotPlayerPlaylistState State, int Score)> scoredStates = states
            .Select(state => (State: state, Score: titleCandidates.Count == 0
                ? (state.IsPreferredSessionFile ? 60 : 50)
                : titleCandidates.Max(title => CalculateMatchScore(title, state.CurrentMediaPath))))
            .Where(item => item.Score > 0)
            .OrderByDescending(item => item.Score)
            .ThenByDescending(item => item.State.IsPreferredSessionFile)
            .ThenByDescending(item => item.State.LastWriteTimeUtc)
            .ToList();

        if (scoredStates.Count == 0) {
            details = $"已读取 {states.Count} 个 PotPlayer 状态文件，但没有任何 `playname` 与窗口标题匹配。";
            return false;
        }

        int bestScore = scoredStates[0].Score;
        PotPlayerPlaylistState bestState = scoredStates[0].State;

        if (bestScore < 70 && titleCandidates.Count > 0) {
            details = $"PotPlayer 状态文件匹配分数过低（最高 {bestScore}），最佳来源: {Path.GetFileName(bestState.PlaylistFilePath)}。";
            return false;
        }

        List<PotPlayerPlaylistState> tiedStates = scoredStates
            .Where(item => item.Score == bestScore)
            .Select(item => item.State)
            .ToList();

        PotPlayerPlaylistState? uniqueState = TryResolveUniqueStateByPath(tiedStates);
        if (uniqueState == null) {
            details = $"PotPlayer 状态文件匹配结果不唯一，候选来源: {string.Join(" ; ", tiedStates.Select(item => Path.GetFileName(item.PlaylistFilePath)))}";
            return false;
        }

        currentMediaPath = uniqueState.CurrentMediaPath;
        details = $"通过 PotPlayer 状态文件命中: {Path.GetFileName(uniqueState.PlaylistFilePath)} -> {uniqueState.CurrentMediaPath}";
        return true;
    }

    /// <summary>
    /// 当多个状态文件得分相同，只有它们指向同一路径时才认定结果唯一。
    /// </summary>
    private static PotPlayerPlaylistState? TryResolveUniqueStateByPath(List<PotPlayerPlaylistState> states) {
        List<string> paths = states
            .Select(item => item.CurrentMediaPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (paths.Count != 1) {
            return null;
        }

        return states
            .OrderByDescending(item => item.IsPreferredSessionFile)
            .ThenByDescending(item => item.LastWriteTimeUtc)
            .First();
    }

    /// <summary>
    /// 枚举 PotPlayer 在 AppData 下维护的播放状态文件。
    /// </summary>
    private static IEnumerable<PotPlayerPlaylistState> ReadPotPlayerPlaylistStates(string processName) {
        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        List<string> playlistDirectories = new();

        string directPlaylistDirectory = Path.Combine(appDataPath, processName, "Playlist");
        if (Directory.Exists(directPlaylistDirectory)) {
            playlistDirectories.Add(directPlaylistDirectory);
        }

        foreach (string potDirectory in Directory
                     .EnumerateDirectories(appDataPath, "PotPlayer*", SearchOption.TopDirectoryOnly)
                     .Where(path => Directory.Exists(Path.Combine(path, "Playlist")))) {
            string playlistDirectory = Path.Combine(potDirectory, "Playlist");
            if (!playlistDirectories.Contains(playlistDirectory, StringComparer.OrdinalIgnoreCase)) {
                playlistDirectories.Add(playlistDirectory);
            }
        }

        string preferredSessionFileName = $"{processName}.dpl";
        foreach (string playlistDirectory in playlistDirectories) {
            foreach (string playlistFilePath in Directory.EnumerateFiles(playlistDirectory, "*.dpl", SearchOption.TopDirectoryOnly)) {
                string? currentPlayname = ReadDplPlayname(playlistFilePath);
                if (string.IsNullOrWhiteSpace(currentPlayname)) {
                    continue;
                }

                yield return new PotPlayerPlaylistState {
                    PlaylistFilePath = playlistFilePath,
                    CurrentMediaPath = currentPlayname.Trim(),
                    LastWriteTimeUtc = File.GetLastWriteTimeUtc(playlistFilePath),
                    IsPreferredSessionFile = string.Equals(Path.GetFileName(playlistFilePath), preferredSessionFileName, StringComparison.OrdinalIgnoreCase),
                };
            }
        }
    }

    /// <summary>
    /// 从 DPL 文件中读取 playname 字段，即当前播放文件的完整路径。
    /// </summary>
    private static string? ReadDplPlayname(string playlistFilePath) {
        try {
            foreach (string rawLine in File.ReadLines(playlistFilePath)) {
                string line = rawLine.Trim();
                if (!line.StartsWith("playname=", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                string value = line[(line.IndexOf('=') + 1)..].Trim();
                return string.IsNullOrWhiteSpace(value) ? null : value;
            }
        } catch {
            return null;
        }

        return null;
    }

    /// <summary>
    /// 批量对候选媒体路径进行标题匹配评分。
    /// </summary>
    private static List<(string Path, int Score)> ScoreCandidates(IEnumerable<string> candidates, IEnumerable<string> titleCandidates) {
        List<string> titleCandidateList = titleCandidates
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (titleCandidateList.Count == 0) {
            return new List<(string Path, int Score)>();
        }

        return candidates
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(path => (Path: path, Score: titleCandidateList.Max(title => CalculateMatchScore(title, path))))
            .Where(item => item.Score > 0)
            .OrderByDescending(item => item.Score)
            .ToList();
    }

    /// <summary>
    /// 从评分结果中挑选唯一且达标的最佳候选路径。
    /// </summary>
    private static bool TryPickBestCandidate(
        List<(string Path, int Score)> scoredCandidates,
        int minScore,
        out string? currentMediaPath,
        out string details) {
        currentMediaPath = null;

        if (scoredCandidates.Count == 0) {
            details = "没有任何候选命中标题。";
            return false;
        }

        int bestScore = scoredCandidates[0].Score;
        if (bestScore < minScore) {
            details = $"匹配分数过低（最高 {bestScore}）。";
            return false;
        }

        List<string> bestPaths = scoredCandidates
            .Where(item => item.Score == bestScore)
            .Select(item => item.Path)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (bestPaths.Count != 1) {
            details = $"匹配结果不唯一（分数 {bestScore}），候选: {string.Join(" ; ", bestPaths.Select(Path.GetFileName))}";
            return false;
        }

        currentMediaPath = bestPaths[0];
        details = $"最佳匹配分数 {bestScore}，命中: {Path.GetFileName(currentMediaPath)}";
        return true;
    }

    /// <summary>
    /// 当命令行只有一个初始媒体文件时，扫描同目录媒体文件做兜底匹配。
    /// </summary>
    private static bool TryResolveFromSiblingFiles(
        string primaryCandidate,
        List<string> titleCandidates,
        out string? currentMediaPath,
        out string details) {
        currentMediaPath = null;
        details = string.Empty;

        string? directory = Path.GetDirectoryName(primaryCandidate);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory)) {
            details = "无法扫描初始媒体所在目录。";
            return false;
        }

        string primaryExtension = Path.GetExtension(primaryCandidate);
        IEnumerable<string> siblingFiles;
        try {
            siblingFiles = Directory.EnumerateFiles(directory)
                .Where(path => MediaExtensions.Contains(Path.GetExtension(path)))
                .ToList();
        } catch (Exception ex) {
            details = $"扫描同目录媒体失败: {ex.Message}";
            return false;
        }

        List<(string Path, int Score)> scoredSiblingFiles = ScoreCandidates(siblingFiles, titleCandidates);
        if (scoredSiblingFiles.Count == 0) {
            details = $"同目录 {directory} 中未找到标题匹配的媒体文件。";
            return false;
        }

        List<(string Path, int Score)> sameExtensionCandidates = scoredSiblingFiles
            .Where(item => string.Equals(Path.GetExtension(item.Path), primaryExtension, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (TryPickBestCandidate(sameExtensionCandidates, 70, out currentMediaPath, out string sameExtensionDetails)) {
            details = $"通过同目录同扩展名扫描命中。{sameExtensionDetails}";
            return true;
        }

        if (TryPickBestCandidate(scoredSiblingFiles, 85, out currentMediaPath, out string siblingBestDetails)) {
            details = $"通过同目录媒体扫描命中。{siblingBestDetails}";
            return true;
        }

        details = $"已扫描同目录媒体，但未得到唯一高分结果。{sameExtensionDetails} {siblingBestDetails}".Trim();
        return false;
    }

    /// <summary>
    /// 从 PotPlayer 原始命令行中提取媒体候选。
    /// 支持直接媒体文件和播放列表文件两种形式。
    /// </summary>
    private static List<string> CollectPotPlayerMediaCandidates(string originalCommandLine) {
        List<string> arguments = CommandLineUtils.ParseArguments(originalCommandLine);
        HashSet<string> candidates = new(StringComparer.OrdinalIgnoreCase);

        foreach (string argument in arguments.Skip(1)) {
            string token = argument.Trim();
            if (string.IsNullOrWhiteSpace(token) || IsOptionToken(token)) {
                continue;
            }

            if (IsPlaylistFile(token)) {
                foreach (string playlistItem in ReadPlaylistEntries(token)) {
                    AddCandidatePath(candidates, playlistItem, Path.GetDirectoryName(token));
                }

                continue;
            }

            if (LooksLikeMediaPath(token)) {
                AddCandidatePath(candidates, token);
            }
        }

        return candidates.ToList();
    }

    /// <summary>
    /// 基于窗口标题构造多个可比较的标题候选，以提升匹配容错率。
    /// </summary>
    private static List<string> BuildPotPlayerTitleCandidates(string windowTitle) {
        HashSet<string> candidates = new(StringComparer.OrdinalIgnoreCase);
        string baseTitle = StripPotPlayerSuffix(windowTitle);
        AddTitleCandidate(candidates, baseTitle);

        string withoutLeadingMetadata = RemoveLeadingMetadata(baseTitle);
        AddTitleCandidate(candidates, withoutLeadingMetadata);

        string withoutTrailingMetadata = RemoveTrailingMetadata(baseTitle);
        AddTitleCandidate(candidates, withoutTrailingMetadata);
        AddTitleCandidate(candidates, RemoveTrailingMetadata(withoutLeadingMetadata));
        AddTitleCandidate(candidates, RemoveKnownNoise(withoutTrailingMetadata));

        foreach (string separator in new[] { " | ", " • ", " · ", " — ", " – " }) {
            int index = baseTitle.IndexOf(separator, StringComparison.Ordinal);
            if (index > 0) {
                AddTitleCandidate(candidates, baseTitle[..index]);
            }
        }

        return candidates.ToList();
    }

    /// <summary>
    /// 对标题片段做轻量清洗后加入候选集合。
    /// </summary>
    private static void AddTitleCandidate(ISet<string> candidates, string value) {
        string normalized = value.Trim().Trim('"', '\'', '-', '_', '·', '•');
        if (!string.IsNullOrWhiteSpace(normalized)) {
            candidates.Add(normalized);
        }
    }

    /// <summary>
    /// 去掉 PotPlayer 窗口标题尾部的播放器后缀，只保留媒体标题主体。
    /// </summary>
    private static string StripPotPlayerSuffix(string windowTitle) {
        string title = windowTitle?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(title)) {
            return string.Empty;
        }

        foreach (string suffix in PotPlayerSuffixes) {
            if (title.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) {
                return title[..^suffix.Length].Trim();
            }
        }

        int lastSeparatorIndex = title.LastIndexOf(" - ", StringComparison.Ordinal);
        if (lastSeparatorIndex > 0) {
            string suffix = title[(lastSeparatorIndex + 3)..];
            if (suffix.Contains("potplayer", StringComparison.OrdinalIgnoreCase)) {
                return title[..lastSeparatorIndex].Trim();
            }
        }

        return title;
    }

    /// <summary>
    /// 去掉标题开头的元信息。
    /// </summary>
    private static string RemoveLeadingMetadata(string title) {
        string result = title.Trim();
        while (LeadingBracketMetadataRegex.IsMatch(result)) {
            result = LeadingBracketMetadataRegex.Replace(result, string.Empty, 1).Trim();
        }

        return result;
    }

    /// <summary>
    /// 去掉标题结尾的元信息。
    /// </summary>
    private static string RemoveTrailingMetadata(string title) {
        string result = title.Trim();
        while (TrailingBracketMetadataRegex.IsMatch(result)) {
            result = TrailingBracketMetadataRegex.Replace(result, string.Empty, 1).Trim();
        }

        return result;
    }

    /// <summary>
    /// 去掉标题里的常见画质、编码和播放时间等噪声文本。
    /// </summary>
    private static string RemoveKnownNoise(string title) {
        string result = title.Trim();
        result = ResolutionNoiseRegex.Replace(result, string.Empty).Trim();
        result = TimeNoiseRegex.Replace(result, string.Empty).Trim();
        return result;
    }

    /// <summary>
    /// 计算标题候选与文件路径之间的相似度分数。
    /// 分数越高，越可能是当前真实播放文件。
    /// </summary>
    private static int CalculateMatchScore(string titleCandidate, string path) {
        string normalizedTitle = NormalizeForComparison(titleCandidate);
        string normalizedFileName = NormalizeForComparison(Path.GetFileName(path));
        string normalizedFileStem = NormalizeForComparison(Path.GetFileNameWithoutExtension(path));

        if (string.IsNullOrWhiteSpace(normalizedTitle) ||
            (string.IsNullOrWhiteSpace(normalizedFileName) && string.IsNullOrWhiteSpace(normalizedFileStem))) {
            return 0;
        }

        if (normalizedTitle == normalizedFileName) {
            return 100;
        }

        if (normalizedTitle == normalizedFileStem) {
            return 95;
        }

        if (normalizedTitle.Contains(normalizedFileName) && normalizedFileName.Length >= 4) {
            return 90;
        }

        if (normalizedTitle.Contains(normalizedFileStem) && normalizedFileStem.Length >= 4) {
            return 88;
        }

        if (normalizedFileName.Contains(normalizedTitle) && normalizedTitle.Length >= 6) {
            return 80;
        }

        if (normalizedFileStem.Contains(normalizedTitle) && normalizedTitle.Length >= 6) {
            return 78;
        }

        int commonPrefixLength = GetCommonPrefixLength(normalizedTitle, normalizedFileStem);
        if (commonPrefixLength >= 8) {
            return 72;
        }

        return 0;
    }

    /// <summary>
    /// 计算两个字符串的公共前缀长度，作为较弱的相似度信号。
    /// </summary>
    private static int GetCommonPrefixLength(string left, string right) {
        int length = Math.Min(left.Length, right.Length);
        int index = 0;
        while (index < length && left[index] == right[index]) {
            index++;
        }

        return index;
    }

    /// <summary>
    /// 将候选路径标准化后加入集合，避免重复和无效路径。
    /// </summary>
    private static void AddCandidatePath(ISet<string> candidates, string candidatePath, string? baseDirectory = null) {
        string? normalizedPath = TryNormalizeCandidatePath(candidatePath, baseDirectory);
        if (!string.IsNullOrWhiteSpace(normalizedPath)) {
            candidates.Add(normalizedPath);
        }
    }

    /// <summary>
    /// 规范化候选路径：去掉引号、处理相对路径并过滤网络地址。
    /// </summary>
    private static string? TryNormalizeCandidatePath(string candidatePath, string? baseDirectory = null) {
        string normalizedPath = candidatePath.Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(normalizedPath)) {
            return null;
        }

        if (normalizedPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            normalizedPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
            return null;
        }

        try {
            if (!Path.IsPathRooted(normalizedPath) && !string.IsNullOrWhiteSpace(baseDirectory)) {
                return Path.GetFullPath(Path.Combine(baseDirectory, normalizedPath));
            }

            return Path.IsPathRooted(normalizedPath)
                ? Path.GetFullPath(normalizedPath)
                : normalizedPath;
        } catch {
            return null;
        }
    }

    /// <summary>
    /// 轻量读取播放列表文件中的媒体项。
    /// </summary>
    private static IEnumerable<string> ReadPlaylistEntries(string playlistPath) {
        string? normalizedPlaylistPath = TryNormalizeCandidatePath(playlistPath);
        if (string.IsNullOrWhiteSpace(normalizedPlaylistPath) || !File.Exists(normalizedPlaylistPath)) {
            yield break;
        }

        string playlistDirectory = Path.GetDirectoryName(normalizedPlaylistPath) ?? Directory.GetCurrentDirectory();

        foreach (string rawLine in File.ReadLines(normalizedPlaylistPath)) {
            string line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith('#') || line.StartsWith(';')) {
                continue;
            }

            if (line.StartsWith("File", StringComparison.OrdinalIgnoreCase) && line.Contains('=')) {
                line = line[(line.IndexOf('=') + 1)..].Trim();
            } else if (line.Contains("*file*", StringComparison.OrdinalIgnoreCase)) {
                int markerIndex = line.IndexOf("*file*", StringComparison.OrdinalIgnoreCase);
                line = line[(markerIndex + 6)..].TrimStart('*').Trim();
            } else if (line.StartsWith('[') && line.EndsWith(']')) {
                continue;
            } else if (line.Contains('=') && !Path.HasExtension(line)) {
                continue;
            }

            string? normalizedItemPath = TryNormalizeCandidatePath(line, playlistDirectory);
            if (!string.IsNullOrWhiteSpace(normalizedItemPath)) {
                yield return normalizedItemPath;
            }
        }
    }

    /// <summary>
    /// 判断路径是否为播放列表文件。
    /// </summary>
    private static bool IsPlaylistFile(string path) {
        string extension = Path.GetExtension(path.Trim().Trim('"'));
        return PlaylistExtensions.Contains(extension);
    }

    /// <summary>
    /// 判断一个参数是否像媒体文件路径，只接受已知媒体扩展名。
    /// </summary>
    private static bool LooksLikeMediaPath(string token) {
        string candidate = token.Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(candidate)) {
            return false;
        }

        string extension = Path.GetExtension(candidate);
        return !string.IsNullOrWhiteSpace(extension) && MediaExtensions.Contains(extension);
    }

    /// <summary>
    /// 判断一个参数是否为命令行选项。
    /// </summary>
    private static bool IsOptionToken(string token) {
        return token.StartsWith('-') || token.StartsWith("/");
    }

    /// <summary>
    /// 将文本归一化为便于比较的形式：小写、去变音符号、仅保留主要字符。
    /// </summary>
    private static string NormalizeForComparison(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        string decomposed = value.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
        StringBuilder builder = new();

        foreach (char ch in decomposed) {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(ch);
            if (category == UnicodeCategory.NonSpacingMark) {
                continue;
            }

            if (char.IsLetterOrDigit(ch) || (ch >= 0x4E00 && ch <= 0x9FFF)) {
                builder.Append(ch);
            }
        }

        return builder.ToString();
    }

    /// <summary>
    /// 获取进程主程序完整路径，优先走 MainModule，失败时回退到 Win32 API。
    /// </summary>
    private static string GetProcessFilePath(Process process) {
        try {
            return process.MainModule?.FileName ?? string.Empty;
        } catch {
            try {
                StringBuilder buffer = new(1024);
                int size = buffer.Capacity;
                return QueryFullProcessImageName(process.Handle, 0, buffer, out size)
                    ? buffer.ToString()
                    : string.Empty;
            } catch {
                return string.Empty;
            }
        }
    }

    /// <summary>
    /// 通过 WMI 查询进程完整命令行。
    /// </summary>
    private static string GetCommandLineArgs(Process process) {
        try {
            using System.Management.ManagementObjectSearcher searcher =
                new($"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}");

            foreach (System.Management.ManagementObject obj in searcher.Get()) {
                return obj["CommandLine"]?.ToString() ?? string.Empty;
            }
        } catch {
            return string.Empty;
        }

        return string.Empty;
    }
}






