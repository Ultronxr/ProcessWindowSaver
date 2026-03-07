using System.Runtime.Versioning;
using ProcessWindowSaver;

[assembly: SupportedOSPlatform("windows")]

string[] commandList = { "0", "1", "2", "3", "4" };
Console.WriteLine("请选择功能：\n[0] 退出\n[1] 保存窗口状态\n[2] 保存指定进程的窗口状态\n[3] 保存 PotPlayer 窗口状态\n[4] 恢复窗口状态");
string? line = Console.ReadLine()?.Trim();

while (string.IsNullOrEmpty(line) || !commandList.Contains(line)) {
    Console.WriteLine("请输入正确的选项编号！");
    line = Console.ReadLine()?.Trim();
}

switch (line) {
    case "0": {
        break;
    }
    case "1": {
        Saver.Save();
        break;
    }
    case "2": {
        Console.WriteLine("请输入进程名称进行筛选（白名单，不区分大小写，多个进程名称之间使用英文逗号分隔）：");
        line = Console.ReadLine()?.Trim();
        Saver.Save(line);
        break;
    }
    case "3": {
        Saver.Save("PotPlayer");
        break;
    }
    case "4": {
        Starter.Start();
        break;
    }
}
