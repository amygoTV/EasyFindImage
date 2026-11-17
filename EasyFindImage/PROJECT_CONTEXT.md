# EasyFindImage — 项目上下文

> 本文件是项目的"记忆"。任何新会话先读这份文件，再读代码。
> 最后更新：2026-11-11

## 一、项目定位

**EasyFindImage** 是一个 Windows 桌面小工具（.NET Framework 4.8 WinForms），面向视觉检测产线的数据整理场景：

1. **表格汇总**：把一个文件夹下的所有表格（CSV/XLS/XLSX）合并成一张表
2. **数据筛选/清理**：按条件筛出需要的行、剔除无效行
3. **按 SN 找图**：从表格或 TXT 读取 SN，在磁盘上找到对应 SN 的图片文件夹并复制到指定路径
4. **一键电脑设置**：把产线 IPC 上常见的系统配置项做成一键勾选执行

## 二、技术栈与工程

| 项 | 值 |
|---|---|
| 框架 | .NET Framework **4.8**（`net48`） |
| 工程类型 | SDK-style WinForms（`Microsoft.NET.Sdk.WindowsDesktop`） |
| NuGet 包 | **无**（零第三方依赖） |
| 额外引用 | 仅 `System.Management`（WMI：还原点、网络适配器枚举） |
| 图标 | `app.ico`（256×256 PNG-based）+ 代码绘制的运行时窗体图标 |
| 构建 | `dotnet build EasyFindImage/EasyFindImage.csproj` |
| 发布 | `dotnet publish EasyFindImage/EasyFindImage.csproj -c Release -o ./publish` |

> `.NET Framework 4.8` 依赖系统已装运行时，发布包**不含**运行时 DLL，只需 `exe` + `config` + `ico`。

### 文件职责

| 文件 | 职责 |
|---|---|
| `Program.cs` | 入口，`Application.Run(new MainForm())` |
| `MainForm.cs` | 主窗体 UI 构建（三个 Tab 的控件创建）+ 图标绘制 |
| `MainForm.Actions.cs` | Tab1/Tab2 的业务逻辑 + 配置加载/保存 |
| `SettingsPage.cs` | Tab3「一键电脑设置」的 UI + 全部设置项实现方法 |
| `ConfigHelper.cs` | `AppConfig` 实体 + XML 序列化读写 `config.xml` |
| `Properties/AssemblyInfo.cs` | 程序集属性（标题/版本等） |

代码风格：**不用 Designer 文件**，全部在 `.cs` 里用代码 new 控件 + 绝对坐标布局。

## 三、三个标签页

### Tab 1「表格汇总与筛选」

```
1. 选择目标文件夹   [路径框] [浏览...] [汇总所有表格]
2. 汇总结果         DataGridView（横向+纵向滚动条）
3. 筛选数据行       [列名▼][关键字][筛选][清除筛选][导出结果]
4. 数据清理条件     [列名▼][关键字][添加条件]
                    [条件列表ListBox] [移除选中][清空条件]
                    [删除重复SN行] [只保留共有列]  [执行清理]
```

**汇总流程**（`BtnSummary_Click`）：
1. 遍历源文件夹（含子目录）所有 `.csv/.xls/.xlsx`
2. 第一张表的结构作为列模板，后续表**按列名匹配**拷贝数据（缺列填 `DBNull.Value`）
3. 自动删除残留表头行（整行每格都等于对应列名）
4. 应用用户自定义清理条件
5. 可选「只保留共有列」→ 删掉所有行都为空的列
6. 刷新列名下拉框（表头自动生成）

**清理条件格式**（关键）：
- 列表中每项形如 `列名:关键字`，多个条件用**英文逗号**分隔
- **同一列多关键字用逗号分隔 = "或"关系**，任一命中即删除该行
- **关键字留空 = 匹配该列为空白**（`SN:` 表示删 SN 列为空的行）
- 例：`SN:,NA,SN,SN1,1` → SN 列为空 / 等于 NA / 等于 SN / 等于 SN1 / 等于 1 的行全删
- ⚠️ 格式陷阱：无冒号的部分会被归到**上一个**列名下；`SN:` 后面必须有内容或留空表示空白

**两个开关可独立执行**：只勾「删除重复SN行」或只加清理条件，都能单独点「执行清理」。

### Tab 2「按 SN 查找图片」

```
1. SN 来源      ○来自表格 ○来自TXT文本  [文件框][浏览...]
2. 图片根目录   [路径框][浏览...]
                查找范围: ○指定目录查找 ○全盘递归查找（从磁盘根目录开始，较慢）
3. 复制到目标路径 [路径框][浏览...]
   [开始查找并复制]
   [进度条]
   [状态标签]
   [日志ListBox]（有滚动条，自动滚到底）
```

**查找规则**（`FindAllSnFoldersRecursive` + `BtnStartCopy_Click`）：
- 遍历所有匹配到的 SN 文件夹 → **按文件夹创建时间排序，取最早的一个**
- 先精确匹配文件夹名 == SN，再模糊匹配文件夹名包含 SN
- 全盘模式只扫固定磁盘根目录；访问被拒的目录静默跳过
- 日志会记录 `[多匹配] SNxxx 找到 N 个文件夹，已选择创建时间最早的：xxx`

### Tab 3「一键电脑设置」

6 个 GroupBox，含约 30 个勾选项，分 6 组：

| 组 | 内容 |
|---|---|
| 可选设置项目（PPTX） | 显示扩展名/隐藏文件、关休眠、清理 TEMP、桌面图标、USB 电源→Disabled、PCIe 电源→Disabled、仅留 eBUS+IPv4、Receive Buffers→最高值、Interrupt Moderation→Extreme、BIOS PCIe Speed→Auto（仅提醒） |
| 新电脑必做 7 项 | Windows 更新、创建还原点、启用 Defender、卸载预装软件、高性能电源、UAC、创建恢复盘 |
| Win10/11 性能优化 | 关 SysMain、关 Windows 搜索、关遥测、关后台应用、禁 OneDrive 自启、视觉效果最佳性能、关技巧建议、关锁屏聚焦 |
| 隐私选项 | 关闭广告 ID、位置服务、诊断数据、在线语音识别 |
| 系统启动设置 | 关闭快速启动 |
| 一键执行 | 批量应用 + 刷新状态 + 日志 |

- 需要管理员权限，不满足时弹窗询问是否以管理员重启（`RestartAsAdmin`）
- 本页有 `tab.AutoScroll = true`，内容超一屏时可滚动

## 四、⚠️ 用户已明确拍板的设计决策（不要擅自推翻）

### 1. 控件样式 = Windows 默认样式
用户**两次否决**了自定义主题：
- 第一次：要求"简洁素雅"，做了钢蓝色系 + 微软雅黑 + 扁平化按钮 + 自绘渐变进度条 + 自绘彩色 Tab 标题
- 用户反馈：**"太丑了，文字都看不见了"** → **"所有控件都恢复默认的形状和颜色吧"**

**结论：控件一律用系统默认外观，不要再引入 `BackColor`/`ForeColor`/`FlatStyle`/自绘。**
（现存唯一的自定义视觉 = 程序图标，用户明确要保留）

### 2. 图标必须保留
- 窗体/任务栏图标：代码在 `MainForm.CreateAppIcon()` 画（蓝色渐变圆 + 白字 "E"），存 `_appIconBitmap` 字段防止被 GC 回收，`OnFormClosing` 里 Dispose
- exe 文件图标：`app.ico` + csproj 的 `<ApplicationIcon>`

### 3. 配置文件
- 位置：**程序启动目录**（exe 同级）的 `config.xml`（不是 AppData）
- **路径类不保存**：源文件夹、SN来源文件、图片根目录、目标路径 **都不写入配置**
- 只保存：筛选列名/关键字、清理条件列表、各复选框开关状态
- 保存时机：`FormClosing`；加载时机：`Load`

### 4. 其他已确认行为
- 汇总时**按列名匹配**合并，不要按位置
- 同一 SN 多个文件夹 → 取**创建时间最早**的
- Excel 读取依赖 OLEDB 提供程序（ACE/JET），未装则该文件被跳过并记日志

## 五、踩过的坑（避免重蹈）

1. **构建失败 `MSB3027 文件被占用`**：旧的 `EasyFindImage.exe` 还在运行会锁住输出文件。
   → 构建前先 `Get-Process EasyFindImage -ErrorAction SilentlyContinue | Stop-Process -Force`

2. **出现两个「一键电脑设置」空白页**：`SettingsPage.BuildSettingsTab` 里自己 `new TabPage` 并 `tabControl.TabPages.Add`，而 MainForm 已经建过一个 → 重复。
   → 正确做法：`BuildSettingsTab(TabPage tab)` 只往传入的 `tab` 上 Add 控件，不新建 TabPage。

3. **`输入数组长度大于此表中的列数`**：不同表格列数不一致时 `Rows.Add(row.ItemArray)` 报错。
   → 已改为按列名构造 `object[]` 再 Add。

4. **模型读不了 PDF/PPTX**：`read` 工具报"does not support pdf input"。PPTX 可用临时 C# 小工具读：`ZipFile.OpenRead` + `Regex.Replace(xml,"<[^>]+>"," ")`。
   → 临时控制台项目要用 `net8.0`（本机只有 net8/net9 运行时，net6.0 编得过但跑不起来）。

5. **`CreateCheckBox` 二义性**：两个 partial 类里都有同名方法，签名第 5 参数一个是 `bool` 一个是 `int` → 调用报 CS0121。现在工厂方法已全部删除，直接 `new CheckBox { ... }`。

## 六、已知限制

- Excel 读取需系统装 Microsoft Access Database Engine（ACE.OLEDB），否则 `.xls/.xlsx` 读取失败被跳过
- `SetReceiveBuffersMax` / `SetInterruptModerationExtreme` 靠猜网卡注册表参数名，不同网卡未必生效
- `powercfg` 的 USB/PCIe GUID 是近似值，不一定精确命中目标设置项
- 清理条件解析无冒号片段会归到上一列，条件写错会静默不生效
- CSV 用 `Encoding.Default`（ANSI）读取，UTF-8 BOM 文件可能乱码
