# ProcessWindowSaver

一个 Windows 桌面窗口状态保存/恢复工具。

## 功能简介

- 扫描当前正在运行、且具有可视主窗口的进程
- 保存应用路径、命令行参数、窗口位置、窗口大小、窗口标题等信息到 JSON 文件
- 读取保存的快照文件，重新启动应用并尽量恢复窗口状态
- 针对 PotPlayer 增强当前播放媒体识别与恢复

> 说明：微软 [PowerToys](https://github.com/microsoft/PowerToys) 的 [工作区](https://learn.microsoft.com/zh-cn/windows/powertoys/workspaces) 提供了相近能力；本项目是面向个人使用场景的轻量实现。

## 运行环境

### 用户运行环境

- 操作系统：Windows 10 / Windows 11
- 架构：建议 `x64`
- 运行时：如果使用默认发布方式，需要安装 `.NET 10 Runtime`
- 权限：部分进程信息读取依赖 WMI / 进程访问权限，遇到受保护进程时可能会跳过并输出提示

### 开发环境

- SDK：`.NET SDK 10.0.103`（由 `global.json` 固定）
- IDE：JetBrains Rider / Visual Studio / VS Code 均可
- Git：用于基于 tag 自动生成发布版本号与文件名

## 本地开发

### 还原与构建

```powershell
dotnet restore .\ProcessWindowSaver.sln
dotnet build .\ProcessWindowSaver.sln -c Debug
```

### 本地运行

```powershell
dotnet run --project .\ProcessWindowSaver.csproj
```

## 发布与版本约定

### 版本来源

项目发布时会读取当前 `HEAD` 所在的 **精确 Git tag**。

支持的 tag 格式：

- `v1.0.0`
- `v1.0.0.1`

如果当前提交没有精确 tag，或者 tag 格式不符合上述规则，`dotnet publish` 会直接失败。

### 版本号映射规则

假设当前 `HEAD` 的 tag 是 `v1.0.0`：

- 文件版本 `FileVersion` = `1.0.0`
- 产品版本 `ProductVersion` = `1.0.0+<commit-hash>`
- 发布后的 exe 文件名 = `ProcessWindowSaver-v1.0.0.exe`

其中：

- `FileVersion` 使用纯数字版本，便于 Windows 文件属性展示
- `ProductVersion` 保留 SDK 注入的源代码提交哈希，便于追踪具体构建来源

### 本地发布

先确保当前提交已经打上符合规则的 tag，例如：

```powershell
git tag v1.0.0
git push origin v1.0.0
```

然后执行：

```powershell
dotnet publish .\ProcessWindowSaver.csproj -c Release
```

默认发布输出目录：

```text
bin\Release\net10.0-windows\publish\
```

如果 tag 为 `v1.0.0`，则目录中会生成：

```text
ProcessWindowSaver-v1.0.0.exe
```

### 常见发布失败原因

- 当前 `HEAD` 没有精确 tag
- tag 不是 `v<major>.<minor>.<patch>` 或 `v<major>.<minor>.<patch>.<revision>` 格式
- 本机未安装 `.NET 10 SDK`
- 本机 `git` 不可用，导致 MSBuild 无法读取 tag

## GitHub Actions

仓库内置了 Windows CI：

- 普通分支推送 / PR：执行 `restore` 和 `build`
- tag 推送：额外执行 `publish`
- tag 发布时会检查版本化后的 exe 是否存在，并上传发布产物

当前工作流文件：

- `.github/workflows/build.yml`

推荐发布流程：

```powershell
git add .
git commit -m "release: prepare v1.0.0"
git tag v1.0.0
git push origin main --tags
```

推送 tag 后，GitHub Actions 会自动生成对应版本名的发布产物。

## 项目结构

```text
Model/   数据模型
Util/    工具类
Program.cs  程序入口
Saver.cs    保存窗口状态
Starter.cs  恢复窗口状态
```
