# ProcessWindowSaver Windows进程窗口保存工具

本程序的核心功能如下：

1. 获取当前运行中的**有可视化窗口的**进程信息（应用程序路径、命令行参数、窗口位置、窗口大小、窗口标题等），导出到文件保存；
2. 读取导出保存的文件，恢复进程运行（应用程序路径、命令行参数、窗口位置、窗口大小）。

注：微软 [PowerToys](https://github.com/microsoft/PowerToys) 的 [工作区](https://learn.microsoft.com/zh-cn/windows/powertoys/workspaces) 有类似功能，本程序是对于我个人需求的针对性开发。
