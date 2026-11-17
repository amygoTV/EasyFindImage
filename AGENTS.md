# AGENTS.md

## 必读
开始任何工作前，先读 `EasyFindImage/PROJECT_CONTEXT.md`（项目结构、功能细节、已知限制、历史坑）。
本文件只放**硬性约束**。

## 项目
`EasyFindImage` — .NET Framework 4.8 WinForms 工具：表格汇总/筛选/清理、按 SN 找图复制、一键电脑设置。
代码风格：**不用 Designer 文件**，全部代码 new 控件 + 绝对坐标布局；零 NuGet 依赖。

## 硬性约束（用户已明确拍板，禁止擅自推翻）

1. **控件一律用 Windows 默认外观。** 不要加 `BackColor`/`ForeColor`/`FlatStyle`，不要自绘控件或 Tab 标题。
   用户已两次否决自定义主题（"太丑了，文字都看不见了" → "所有控件都恢复默认的形状和颜色吧"）。
   唯一保留的自定义视觉是**程序图标**（`MainForm.CreateAppIcon()` + `app.ico`），必须保留。

2. **配置文件位置**：程序启动目录（exe 同级）的 `config.xml`。
   **路径类一律不保存**（源文件夹、SN来源文件、图片根目录、目标路径）；只保存筛选条件、清理条件列表、复选框开关。

3. **合并表格按列名匹配**，不要按列位置（各文件列数/顺序可能不同）。

4. **清理条件格式**：`列名:关键字`，逗号分隔；同一列多关键字 = "或"关系；**关键字留空 = 该列为空白**。

5. **同一 SN 多个文件夹 → 取创建时间最早的**。

6. **筛选和清理必须能各自独立执行**（只勾重复SN 或 只加条件，都能单独点执行）。

## 构建注意
构建前先结束占用进程，否则报 `MSB3027 文件被占用`：
```powershell
Get-Process EasyFindImage -ErrorAction SilentlyContinue | Stop-Process -Force
dotnet build EasyFindImage/EasyFindImage.csproj
```

## 沟通
用中文回复。改动后务必跑一次 build 确认 0 错误 0 警告。
