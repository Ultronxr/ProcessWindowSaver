namespace ProcessWindowSaver.Model;

// 进程信息类
public class ProcessInfo {
    public string ProcessName { get; set; }
    public string FilePath { get; set; }
    public string CommandLine { get; set; }
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public string WindowTitle { get; set; }
}