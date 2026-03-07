using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Text.RegularExpressions;
using ProcessWindowSaver.Model;
using ProcessWindowSaver.Util;

namespace ProcessWindowSaver;

[SupportedOSPlatform("windows")]
public class Saver {
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("kernel32.dll")]
    private static extern bool QueryFullProcessImageName(IntPtr hProcess, int dwFlags, StringBuilder lpExeName, out int size);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;

        public int Width => Right - Left;
        public int Height => Bottom - Top;
    }

    private sealed class PotPlayerPlaylistState {
        public string PlaylistFilePath { get; init; } = string.Empty;
        public string CurrentMediaPath { get; init; } = string.Empty;
        public DateTime LastWriteTimeUtc { get; init; }
        public bool IsPreferredSessionFile { get; init; }
    }

    private static readonly HashSet<string> PlaylistExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".m3u", ".m3u8", ".pls", ".dpl", ".asx", ".wax", ".wvx", ".wpl", ".cue"
    };

    private static readonly HashSet<string> MediaExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm", ".m4v", ".ts", ".m2ts", ".mpg", ".mpeg",
        ".mp3", ".flac", ".aac", ".wav", ".ogg", ".m4a", ".ape", ".wma"
    };

    private static readonly string[] PotPlayerSuffixes = {
        " - PotPlayer",
        " - PotPlayerMini",
        " - PotPlayerMini64",
        " - PotPlayer64",
        " - Daum PotPlayer"
    };

    private static readonly Regex LeadingBracketMetadataRegex = new(@"^\s*[\[(][^\])]{1,40}[\])]\s*", RegexOptions.Compiled);
    private static readonly Regex TrailingBracketMetadataRegex = new(@"\s*[\[(][^\])]{1,40}[\])]\s*$", RegexOptions.Compiled);
    private static readonly Regex ResolutionNoiseRegex = new(@"\s+(2160p|1080p|720p|540p|480p|x265|x264|hevc|av1|hdr10?\+?|\d{2,3}fps)\b.*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TimeNoiseRegex = new(@"\s+\d{1,2}:\d{2}(?::\d{2})?(\s*/\s*\d{1,2}:\d{2}(?::\d{2})?)?\s*$", RegexOptions.Compiled);

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

    private static bool IsPotPlayerProcess(string processName) {
        return processName.Contains("potplayer", StringComparison.OrdinalIgnoreCase);
    }

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

    private static void AddTitleCandidate(ISet<string> candidates, string value) {
        string normalized = value.Trim().Trim('"', '\'', '-', '_', '·', '•');
        if (!string.IsNullOrWhiteSpace(normalized)) {
            candidates.Add(normalized);
        }
    }

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

    private static string RemoveLeadingMetadata(string title) {
        string result = title.Trim();
        while (LeadingBracketMetadataRegex.IsMatch(result)) {
            result = LeadingBracketMetadataRegex.Replace(result, string.Empty, 1).Trim();
        }

        return result;
    }

    private static string RemoveTrailingMetadata(string title) {
        string result = title.Trim();
        while (TrailingBracketMetadataRegex.IsMatch(result)) {
            result = TrailingBracketMetadataRegex.Replace(result, string.Empty, 1).Trim();
        }

        return result;
    }

    private static string RemoveKnownNoise(string title) {
        string result = title.Trim();
        result = ResolutionNoiseRegex.Replace(result, string.Empty).Trim();
        result = TimeNoiseRegex.Replace(result, string.Empty).Trim();
        return result;
    }

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

    private static int GetCommonPrefixLength(string left, string right) {
        int length = Math.Min(left.Length, right.Length);
        int index = 0;
        while (index < length && left[index] == right[index]) {
            index++;
        }

        return index;
    }

    private static void AddCandidatePath(ISet<string> candidates, string candidatePath, string? baseDirectory = null) {
        string? normalizedPath = TryNormalizeCandidatePath(candidatePath, baseDirectory);
        if (!string.IsNullOrWhiteSpace(normalizedPath)) {
            candidates.Add(normalizedPath);
        }
    }

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

    private static bool IsPlaylistFile(string path) {
        string extension = Path.GetExtension(path.Trim().Trim('"'));
        return PlaylistExtensions.Contains(extension);
    }

    private static bool LooksLikeMediaPath(string token) {
        string candidate = token.Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(candidate)) {
            return false;
        }

        string extension = Path.GetExtension(candidate);
        return !string.IsNullOrWhiteSpace(extension) && MediaExtensions.Contains(extension);
    }

    private static bool IsOptionToken(string token) {
        return token.StartsWith('-') || token.StartsWith("/");
    }

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




