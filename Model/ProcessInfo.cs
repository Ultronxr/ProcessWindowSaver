namespace ProcessWindowSaver.Model;

/// <summary>
/// 单个可恢复窗口的快照信息。
/// 这份模型既承担“保存时的数据载体”，也承担“恢复时的输入参数”。
/// </summary>
public class ProcessInfo {
    /// <summary>
    /// 默认恢复模式。
    /// 表示恢复阶段使用原始进程命令行来重新启动程序。
    /// 适用于绝大多数普通 GUI 程序。
    /// </summary>
    public const string DefaultCommandLineResumeMode = "DefaultCommandLine";

    /// <summary>
    /// PotPlayer 当前媒体恢复模式。
    /// 表示恢复阶段优先使用识别到的当前播放文件路径来启动 PotPlayer，
    /// 而不是使用最初启动进程时的命令行参数。
    /// </summary>
    public const string PotPlayerCurrentMediaResumeMode = "PotPlayerCurrentMedia";

    /// <summary>
    /// 进程名称，例如 PotPlayerMini64、Code、chrome。
    /// </summary>
    public string ProcessName { get; set; } = string.Empty;

    /// <summary>
    /// 进程主程序的完整路径。
    /// 恢复时若原始命令行不可用，会回退使用该路径启动。
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// 进程最初启动时的完整命令行。
    /// 对大多数程序来说，这就是恢复进程时最可靠的输入。
    /// 对 PotPlayer 而言，它只是“启动来源”，不一定代表“当前播放状态”。
    /// </summary>
    public string OriginalCommandLine { get; set; } = string.Empty;

    /// <summary>
    /// 窗口左上角 X 坐标。
    /// </summary>
    public int X { get; set; }

    /// <summary>
    /// 窗口左上角 Y 坐标。
    /// </summary>
    public int Y { get; set; }

    /// <summary>
    /// 窗口宽度。
    /// </summary>
    public int Width { get; set; }

    /// <summary>
    /// 窗口高度。
    /// </summary>
    public int Height { get; set; }

    /// <summary>
    /// 当前主窗口标题。
    /// 保存阶段会用它辅助识别 PotPlayer 正在播放的媒体。
    /// </summary>
    public string WindowTitle { get; set; } = string.Empty;

    /// <summary>
    /// 当前媒体文件的完整路径。
    /// 仅对支持“当前内容状态识别”的应用有意义，目前主要用于 PotPlayer。
    /// </summary>
    public string? CurrentMediaPath { get; set; }

    /// <summary>
    /// 恢复模式。
    /// 取值通常为 <see cref="DefaultCommandLineResumeMode"/> 或 <see cref="PotPlayerCurrentMediaResumeMode"/>。
    /// </summary>
    public string ResumeMode { get; set; } = DefaultCommandLineResumeMode;
}
