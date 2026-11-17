# EasyFindImage

EasyFindImage 是一个基于 .NET Framework 4.8 + WinForms 开发的桌面小工具，面向视觉检测产线场景，主要功能包括：

## 📋 功能概述

### 1. 表格汇总与筛选
- 将目标文件夹内所有表格（CSV/XLS/XLSX）按列结构汇总到一张表中
- 支持按列名 + 关键字筛选数据行
- 支持导出结果为 CSV

### 2. 按 SN 查找图片
- 从表格（自动识别 SN 列）或 TXT（每行一个 SN）提取 SN
- 在指定的图片根目录中查找文件夹名匹配 SN 的目录（支持精确匹配和模糊匹配）
- 支持从磁盘根目录递归全盘查找（适合 SN 文件夹分布在多个盘符的情况）
- 同一个 SN 若找到多个文件夹，自动选择**创建时间最早**的一个
- 整个文件夹复制到可选目标路径，带进度条和详细日志

### 3. 一键电脑设置
包含约 30 项常见电脑设置，分为：
- **常用设置**：显示文件扩展名、显示隐藏文件、关闭休眠、清理临时文件、桌面图标
- **PPTX 建议设置**（基于 issues to improve.pptx）：USB/PCIe 电源选项、网络适配器配置（eBUS+IPv4、Receive Buffers、Interrupt Moderation）
- **新电脑必做 7 项**：Windows 更新、系统还原点、Defender、卸载预装软件、高性能电源、UAC、系统恢复盘
- **Win10/11 系统优化**：SysMain、Windows 搜索、遥测、后台应用、OneDrive 自启、视觉效果、技巧建议、锁屏聚焦
- **隐私选项**：广告 ID、位置服务、诊断数据、在线语音识别
- **系统启动设置**：快速启动

## 🛠 技术栈

| 项 | 说明 |
|---|---|
| 框架 | .NET Framework 4.8 |
| 界面 | Windows Forms (WinForms) |
| 项目类型 | SDK-style Windows Desktop 项目 |
| 外部依赖 | 无 NuGet 包，仅引用 `System.Management`（WMI） |
| 编译器 | .NET SDK |

## 📁 项目结构

```
EasyFindImage/
├── EasyFindImage.csproj        # 项目文件 (.NET 4.8 WinForms)
├── Program.cs                   # 程序入口
├── MainForm.cs                  # 主窗体 UI 布局
├── MainForm.Actions.cs          # 业务逻辑（汇总、筛选、清理、复制、设置实现）
├── SettingsPage.cs              # 一键电脑设置标签页 UI + 实现
├── ConfigHelper.cs              # 配置持久化（XML，保存到启动目录）
├── app.ico                      # 应用程序图标
├── Properties/
│   └── AssemblyInfo.cs          # 程序集信息
├── AGENTS.md                    # AI 协作硬性约束文档
├── PROJECT_CONTEXT.md           # 项目详细上下文文档
├── README.md                    # 本文件
└── .gitignore                   # Git 忽略文件
```

## 🚀 如何运行

### 方式 1：直接运行已编译版本
从 [Releases](https://github.com/amygoTV/EasyFindImage/releases) 下载发布包，解压后双击 `EasyFindImage.exe` 即可运行（需系统已安装 .NET Framework 4.8）。

### 方式 2：从源码编译
```bash
# 克隆仓库
git clone https://github.com/amygoTV/EasyFindImage.git
cd EasyFindImage

# 构建
dotnet build EasyFindImage/EasyFindImage.csproj

# 运行
dotnet run --project EasyFindImage/EasyFindImage.csproj
```

## ⚠️ 注意事项

1. **Excel 读取**：读取 `.xls`/`.xlsx` 使用 OleDb，需要系统安装 Microsoft Access Database Engine（ACE.OLEDB 或 Jet.OLEDB）。未安装则 Excel 表格会被跳过。
2. **管理员权限**：部分电脑设置（关闭休眠、修改注册表、创建还原点等）需要管理员权限，程序会提示以管理员身份重新启动。
3. **配置文件**：程序会在启动目录生成 `config.xml`，保存复选框状态等设置。**文件路径不会保存**，每次打开需要重新选择。
4. **CSV 编码**：默认使用系统 ANSI 编码读取 CSV，若表格为 UTF-8 请修改代码中的编码设置。

## 📄 许可证

本项目仅供个人学习和内部使用，未添加开源许可证。

---

*基于 Cognex VisionPro 视觉检测产线场景开发*
