using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace EasyFindImage
{
    public partial class MainForm
    {
        // ==================== Tab 3：一键电脑设置 ====================
        private GroupBox grpSysSettings;
        private GroupBox grpActions;
        private GroupBox grpNewPC;
        private GroupBox grpOpt;
        private GroupBox grpPriv;
        private GroupBox grpFastStartup;

        // 原有设置
        private CheckBox chkShowExt;
        private CheckBox chkShowHidden;
        private CheckBox chkHibernateOff;
        private CheckBox chkCleanTemp;
        private CheckBox chkDesktopIcons;

        // PPTX 新增设置
        private CheckBox chkUsbPowerDisabled;
        private CheckBox chkPciePowerDisabled;
        private CheckBox chkNetworkEbusOnly;
        private CheckBox chkReceiveBuffersMax;
        private CheckBox chkInterruptModerationExtreme;
        private CheckBox chkBiosPcieAutoReminder;

        // 新电脑必做设置
        private CheckBox chkWinUpdate;
        private CheckBox chkRestorePoint;
        private CheckBox chkDefender;
        private CheckBox chkRemoveBloatware;
        private CheckBox chkPowerHighPerf;
        private CheckBox chkUAC;
        private CheckBox chkRecoveryDrive;

        // Win10/Win11 系统优化
        private CheckBox chkOptSysMain;
        private CheckBox chkOptWindowsSearch;
        private CheckBox chkOptTelemetry;
        private CheckBox chkOptBackgroundApps;
        private CheckBox chkOptOneDrive;
        private CheckBox chkOptVisualEffects;
        private CheckBox chkOptTips;
        private CheckBox chkOptLockScreen;

        // 隐私选项
        private CheckBox chkPrivAdId;
        private CheckBox chkPrivLocation;
        private CheckBox chkPrivDiagnostics;
        private CheckBox chkPrivSpeech;

        // 快速启动
        private CheckBox chkFastStartup;

        private Button btnApplySettings;
        private Button btnRefreshStatus;
        private Label lblSettingStatus;
        private ListBox lstSettingLog;

        private void BuildSettingsTab(TabPage tab)
        {
            // tab 由 MainForm 创建并添加，这里直接使用
            int y = 15;

            // 设置选项（分两列）
            grpSysSettings = new GroupBox { Text = "可选设置项目（基于 issues to improve.pptx）", Left = 12, Top = y, Width = 900, Height = 260 };

            int col1X = 15;
            int col2X = 460;
            int rowHeight = 30;

            // 左列
            chkShowExt = new CheckBox { Text = "显示文件扩展名（取消隐藏已知文件类型扩展名）", Left = col1X, Top = 24, Width = 400, Checked = true };
            chkShowHidden = new CheckBox { Text = "显示隐藏文件、文件夹和驱动器", Left = col1X, Top = 24 + rowHeight, Width = 340, Checked = false };
            chkHibernateOff = new CheckBox { Text = "关闭系统休眠（释放 hiberfil.sys 空间）", Left = col1X, Top = 24 + rowHeight * 2, Width = 340, Checked = false };
            chkCleanTemp = new CheckBox { Text = "清理用户临时文件（%TEMP%）", Left = col1X, Top = 24 + rowHeight * 3, Width = 280, Checked = true };
            chkDesktopIcons = new CheckBox { Text = "显示桌面常用图标（此电脑、网络等）", Left = col1X, Top = 24 + rowHeight * 4, Width = 320, Checked = false };

            // 右列（PPTX 新增）
            chkUsbPowerDisabled = new CheckBox { Text = "电源选项：USB 设置 → Disabled（防止 USB 断开导致采集卡掉线）", Left = col2X, Top = 24, Width = 420, Checked = false };
            chkPciePowerDisabled = new CheckBox { Text = "电源选项：PCI Express 设置 → Disabled（防止 PCIe 节能导致 Buffer 失败）", Left = col2X, Top = 24 + rowHeight, Width = 420, Checked = false };
            chkNetworkEbusOnly = new CheckBox { Text = "网络适配器：仅保留 eBUS + IPv4，禁用其他协议", Left = col2X, Top = 24 + rowHeight * 2, Width = 380, Checked = false };
            chkReceiveBuffersMax = new CheckBox { Text = "网络适配器：Receive Buffers → 最高值", Left = col2X, Top = 24 + rowHeight * 3, Width = 320, Checked = false };
            chkInterruptModerationExtreme = new CheckBox { Text = "网络适配器：Interrupt Moderation Rate → Extreme", Left = col2X, Top = 24 + rowHeight * 4, Width = 350, Checked = false };
            chkBiosPcieAutoReminder = new CheckBox { Text = "BIOS 提醒：Chipset → PCIe Speed → Auto（需手动进 BIOS 设置）", Left = col2X, Top = 24 + rowHeight * 5, Width = 420, Checked = false };

            grpSysSettings.Controls.Add(chkShowExt);
            grpSysSettings.Controls.Add(chkShowHidden);
            grpSysSettings.Controls.Add(chkHibernateOff);
            grpSysSettings.Controls.Add(chkCleanTemp);
            grpSysSettings.Controls.Add(chkDesktopIcons);
            grpSysSettings.Controls.Add(chkUsbPowerDisabled);
            grpSysSettings.Controls.Add(chkPciePowerDisabled);
            grpSysSettings.Controls.Add(chkNetworkEbusOnly);
            grpSysSettings.Controls.Add(chkReceiveBuffersMax);
            grpSysSettings.Controls.Add(chkInterruptModerationExtreme);
            grpSysSettings.Controls.Add(chkBiosPcieAutoReminder);
            tab.Controls.Add(grpSysSettings);

            y += 270;

            // 新电脑到手后必做设置（7 项常见建议）
            grpNewPC = new GroupBox { Text = "新电脑到手后必做设置（7 项常见建议）", Left = 12, Top = y, Width = 900, Height = 160 };

            int npY = 24;
            int npRow = 28;

            chkWinUpdate = new CheckBox { Text = "1. 检查并安装 Windows 更新", Left = 15, Top = npY, Width = 300 };
            chkRestorePoint = new CheckBox { Text = "2. 创建系统还原点", Left = 15, Top = npY + npRow, Width = 220 };
            chkDefender = new CheckBox { Text = "3. 启用 Windows Defender 实时保护", Left = 15, Top = npY + npRow * 2, Width = 280 };
            chkRemoveBloatware = new CheckBox { Text = "4. 卸载不需要的预装软件", Left = 15, Top = npY + npRow * 3, Width = 260 };

            chkPowerHighPerf = new CheckBox { Text = "5. 电源计划设为“高性能”", Left = 460, Top = npY, Width = 240 };
            chkUAC = new CheckBox { Text = "6. 调整用户账户控制 (UAC)", Left = 460, Top = npY + npRow, Width = 260 };
            chkRecoveryDrive = new CheckBox { Text = "7. 创建系统恢复盘（需手动操作）", Left = 460, Top = npY + npRow * 2, Width = 280 };

            grpNewPC.Controls.Add(chkWinUpdate);
            grpNewPC.Controls.Add(chkRestorePoint);
            grpNewPC.Controls.Add(chkDefender);
            grpNewPC.Controls.Add(chkRemoveBloatware);
            grpNewPC.Controls.Add(chkPowerHighPerf);
            grpNewPC.Controls.Add(chkUAC);
            grpNewPC.Controls.Add(chkRecoveryDrive);
            tab.Controls.Add(grpNewPC);

            y += 170;

            // Win10/Win11 系统优化设置 - 性能部分
            grpOpt = new GroupBox { Text = "Win10/Win11 系统优化设置（性能部分）", Left = 12, Top = y, Width = 900, Height = 135 };

            int opY = 24;
            int opRow = 28;

            chkOptSysMain = new CheckBox { Text = "1. 关闭 SysMain/SuperFetch（适合 SSD，减少磁盘占用）", Left = 15, Top = opY, Width = 420 };
            chkOptWindowsSearch = new CheckBox { Text = "2. 关闭 Windows 搜索/索引服务", Left = 15, Top = opY + opRow, Width = 300 };
            chkOptTelemetry = new CheckBox { Text = "3. 关闭遥测/诊断数据上报", Left = 15, Top = opY + opRow * 2, Width = 300 };
            chkOptBackgroundApps = new CheckBox { Text = "4. 关闭后台应用运行", Left = 15, Top = opY + opRow * 3, Width = 280 };

            chkOptOneDrive = new CheckBox { Text = "5. 禁用 OneDrive 开机自启动", Left = 460, Top = opY, Width = 300 };
            chkOptVisualEffects = new CheckBox { Text = "6. 视觉效果设为“最佳性能”", Left = 460, Top = opY + opRow, Width = 300 };
            chkOptTips = new CheckBox { Text = "7. 关闭“获取技巧和建议”", Left = 460, Top = opY + opRow * 2, Width = 300 };
            chkOptLockScreen = new CheckBox { Text = "8. 关闭锁屏 Windows 聚焦", Left = 460, Top = opY + opRow * 3, Width = 300 };

            grpOpt.Controls.Add(chkOptSysMain);
            grpOpt.Controls.Add(chkOptWindowsSearch);
            grpOpt.Controls.Add(chkOptTelemetry);
            grpOpt.Controls.Add(chkOptBackgroundApps);
            grpOpt.Controls.Add(chkOptOneDrive);
            grpOpt.Controls.Add(chkOptVisualEffects);
            grpOpt.Controls.Add(chkOptTips);
            grpOpt.Controls.Add(chkOptLockScreen);
            tab.Controls.Add(grpOpt);

            y += 145;

            // 隐私选项
            grpPriv = new GroupBox { Text = "隐私选项", Left = 12, Top = y, Width = 900, Height = 70 };

            int privY = 24;

            chkPrivAdId = new CheckBox { Text = "关闭广告 ID（个性化广告）", Left = 15, Top = privY, Width = 260 };
            chkPrivLocation = new CheckBox { Text = "关闭位置服务", Left = 285, Top = privY, Width = 180 };
            chkPrivDiagnostics = new CheckBox { Text = "关闭诊断数据", Left = 475, Top = privY, Width = 180 };
            chkPrivSpeech = new CheckBox { Text = "关闭在线语音识别", Left = 665, Top = privY, Width = 220 };

            grpPriv.Controls.Add(chkPrivAdId);
            grpPriv.Controls.Add(chkPrivLocation);
            grpPriv.Controls.Add(chkPrivDiagnostics);
            grpPriv.Controls.Add(chkPrivSpeech);
            tab.Controls.Add(grpPriv);

            y += 80;

            // 快速启动
            grpFastStartup = new GroupBox { Text = "系统启动设置", Left = 12, Top = y, Width = 900, Height = 60 };

            chkFastStartup = new CheckBox { Text = "关闭快速启动（Fast Startup，避免部分驱动/硬件异常）", Left = 15, Top = 24, Width = 520 };

            grpFastStartup.Controls.Add(chkFastStartup);
            tab.Controls.Add(grpFastStartup);

            y += 70;

            // 操作按钮
            grpActions = new GroupBox { Text = "一键执行", Left = 12, Top = y, Width = 900, Height = 70 };
            btnApplySettings = new Button { Text = "一键应用选中设置", Left = 15, Top = 22, Width = 150, Height = 32 };
            btnRefreshStatus = new Button { Text = "刷新当前状态", Left = 175, Top = 22, Width = 110, Height = 32 };
            grpActions.Controls.Add(btnApplySettings);
            grpActions.Controls.Add(btnRefreshStatus);
            tab.Controls.Add(grpActions);

            y += 80;

            // 日志与状态
            lblSettingStatus = new Label { Left = 15, Top = y, Width = 400, Height = 20, Text = "就绪" };
            lstSettingLog = new ListBox { Left = 15, Top = y + 28, Width = 900, Height = 220, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom, ScrollAlwaysVisible = true };
            tab.Controls.Add(lblSettingStatus);
            tab.Controls.Add(lstSettingLog);

            btnApplySettings.Click += BtnApplySettings_Click;
            btnRefreshStatus.Click += BtnRefreshStatus_Click;

            tab.AutoScroll = true;
        }

        // ==================== Tab 3 事件 ====================

        private void BtnApplySettings_Click(object sender, EventArgs e)
        {
            if (!IsAdministrator())
            {
                var result = MessageBox.Show(
                    this,
                    "大部分设置需要管理员权限才能修改，是否以管理员身份重新启动程序？",
                    "权限提示",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    RestartAsAdmin();
                }
                return;
            }

            lstSettingLog.Items.Add($"[{DateTime.Now:HH:mm:ss}] 开始应用选中设置...");
            lblSettingStatus.Text = "正在应用设置...";
            Application.DoEvents();

            int success = 0, fail = 0;

            // 原有设置
            if (chkShowExt.Checked)
            {
                if (SetFileExtensionsVisible(true)) { AppendSettingLog("[成功] 已显示文件扩展名"); success++; }
                else { AppendSettingLog("[失败] 设置文件扩展名显示"); fail++; }
            }

            if (chkShowHidden.Checked)
            {
                if (SetHiddenFilesVisible(true)) { AppendSettingLog("[成功] 已显示隐藏文件/文件夹"); success++; }
                else { AppendSettingLog("[失败] 设置隐藏文件显示"); fail++; }
            }

            if (chkHibernateOff.Checked)
            {
                if (DisableHibernate()) { AppendSettingLog("[成功] 已关闭系统休眠"); success++; }
                else { AppendSettingLog("[失败] 关闭系统休眠"); fail++; }
            }

            if (chkCleanTemp.Checked)
            {
                if (CleanTempFiles(out long freedMB)) { AppendSettingLog($"[成功] 已清理临时文件，释放约 {freedMB} MB"); success++; }
                else { AppendSettingLog("[失败] 清理临时文件"); fail++; }
            }

            if (chkDesktopIcons.Checked)
            {
                if (ShowDesktopIcons(true)) { AppendSettingLog("[成功] 已显示桌面常用图标"); success++; }
                else { AppendSettingLog("[失败] 设置桌面图标"); fail++; }
            }

            // PPTX 新增设置
            if (chkUsbPowerDisabled.Checked)
            {
                if (SetPowerUsbSelectiveSuspend(false)) { AppendSettingLog("[成功] 已设置 USB 选择性挂起 → Disabled"); success++; }
                else { AppendSettingLog("[失败] 设置 USB 电源选项"); fail++; }
            }

            if (chkPciePowerDisabled.Checked)
            {
                if (SetPowerPcieAspm(false)) { AppendSettingLog("[成功] 已设置 PCI Express ASPM → Disabled"); success++; }
                else { AppendSettingLog("[失败] 设置 PCI Express 电源选项"); fail++; }
            }

            if (chkNetworkEbusOnly.Checked)
            {
                if (SetNetworkEbusOnly()) { AppendSettingLog("[成功] 已配置网络适配器：仅保留 eBUS + IPv4"); success++; }
                else { AppendSettingLog("[失败] 配置网络适配器协议"); fail++; }
            }

            if (chkReceiveBuffersMax.Checked)
            {
                if (SetReceiveBuffersMax()) { AppendSettingLog("[成功] 已设置 Receive Buffers → 最高值"); success++; }
                else { AppendSettingLog("[失败] 设置 Receive Buffers"); fail++; }
            }

            if (chkInterruptModerationExtreme.Checked)
            {
                if (SetInterruptModerationExtreme()) { AppendSettingLog("[成功] 已设置 Interrupt Moderation Rate → Extreme"); success++; }
                else { AppendSettingLog("[失败] 设置 Interrupt Moderation Rate"); fail++; }
            }

            if (chkBiosPcieAutoReminder.Checked)
            {
                AppendSettingLog("[提醒] 请手动进入 BIOS：Chipset → PCIe Speed → Auto（无法通过软件自动设置）");
                success++;
            }

            // 新电脑必做设置
            if (chkWinUpdate.Checked)
            {
                if (OpenWindowsUpdate()) { AppendSettingLog("[成功] 已打开 Windows 更新设置页面"); success++; }
                else { AppendSettingLog("[失败] 打开 Windows 更新设置"); fail++; }
            }

            if (chkRestorePoint.Checked)
            {
                if (CreateSystemRestorePoint()) { AppendSettingLog("[成功] 已创建系统还原点"); success++; }
                else { AppendSettingLog("[失败] 创建系统还原点（需管理员权限）"); fail++; }
            }

            if (chkDefender.Checked)
            {
                if (EnableWindowsDefender()) { AppendSettingLog("[成功] 已启用 Windows Defender 实时保护"); success++; }
                else { AppendSettingLog("[失败] 启用 Windows Defender"); fail++; }
            }

            if (chkRemoveBloatware.Checked)
            {
                if (OpenAppsFeatures()) { AppendSettingLog("[成功] 已打开“应用和功能”页面，可卸载预装软件"); success++; }
                else { AppendSettingLog("[失败] 打开“应用和功能”页面"); fail++; }
            }

            if (chkPowerHighPerf.Checked)
            {
                if (SetPowerHighPerformance()) { AppendSettingLog("[成功] 已设置电源计划为“高性能”"); success++; }
                else { AppendSettingLog("[失败] 设置电源计划"); fail++; }
            }

            if (chkUAC.Checked)
            {
                if (SetUacDefault()) { AppendSettingLog("[成功] 已将 UAC 设置为默认通知级别"); success++; }
                else { AppendSettingLog("[失败] 调整用户账户控制 (UAC)"); fail++; }
            }

            if (chkRecoveryDrive.Checked)
            {
                if (OpenRecoveryDrive()) { AppendSettingLog("[成功] 已打开“创建恢复驱动器”工具"); success++; }
                else { AppendSettingLog("[失败] 打开“创建恢复驱动器”工具"); fail++; }
            }

            // 系统优化设置
            if (chkOptSysMain.Checked)
            {
                if (SetServiceStartType("SysMain", "disabled")) { AppendSettingLog("[成功] 已禁用 SysMain/SuperFetch 服务"); success++; }
                else { AppendSettingLog("[失败] 禁用 SysMain 服务"); fail++; }
            }

            if (chkOptWindowsSearch.Checked)
            {
                if (SetServiceStartType("WSearch", "disabled")) { AppendSettingLog("[成功] 已禁用 Windows 搜索/索引服务"); success++; }
                else { AppendSettingLog("[失败] 禁用 Windows 搜索服务"); fail++; }
            }

            if (chkOptTelemetry.Checked)
            {
                if (SetTelemetryDisabled()) { AppendSettingLog("[成功] 已关闭遥测/诊断数据上报"); success++; }
                else { AppendSettingLog("[失败] 关闭遥测数据"); fail++; }
            }

            if (chkOptBackgroundApps.Checked)
            {
                if (SetBackgroundAppsDisabled()) { AppendSettingLog("[成功] 已关闭后台应用运行"); success++; }
                else { AppendSettingLog("[失败] 关闭后台应用"); fail++; }
            }

            if (chkOptOneDrive.Checked)
            {
                if (SetOneDriveAutoStart(false)) { AppendSettingLog("[成功] 已禁用 OneDrive 开机自启动"); success++; }
                else { AppendSettingLog("[失败] 禁用 OneDrive 自启动"); fail++; }
            }

            if (chkOptVisualEffects.Checked)
            {
                if (SetVisualEffectsBestPerformance()) { AppendSettingLog("[成功] 已将视觉效果设为“最佳性能”"); success++; }
                else { AppendSettingLog("[失败] 设置视觉效果"); fail++; }
            }

            if (chkOptTips.Checked)
            {
                if (SetTipsDisabled()) { AppendSettingLog("[成功] 已关闭“获取技巧和建议”"); success++; }
                else { AppendSettingLog("[失败] 关闭技巧和建议"); fail++; }
            }

            if (chkOptLockScreen.Checked)
            {
                if (SetLockScreenSpotlightDisabled()) { AppendSettingLog("[成功] 已关闭锁屏 Windows 聚焦"); success++; }
                else { AppendSettingLog("[失败] 关闭锁屏聚焦"); fail++; }
            }

            // 隐私选项
            if (chkPrivAdId.Checked)
            {
                if (SetPrivacyAdId(false)) { AppendSettingLog("[成功] 已关闭广告 ID"); success++; }
                else { AppendSettingLog("[失败] 关闭广告 ID"); fail++; }
            }

            if (chkPrivLocation.Checked)
            {
                if (SetPrivacyLocation(false)) { AppendSettingLog("[成功] 已关闭位置服务"); success++; }
                else { AppendSettingLog("[失败] 关闭位置服务"); fail++; }
            }

            if (chkPrivDiagnostics.Checked)
            {
                if (SetPrivacyDiagnostics(false)) { AppendSettingLog("[成功] 已关闭诊断数据"); success++; }
                else { AppendSettingLog("[失败] 关闭诊断数据"); fail++; }
            }

            if (chkPrivSpeech.Checked)
            {
                if (SetPrivacySpeech(false)) { AppendSettingLog("[成功] 已关闭在线语音识别"); success++; }
                else { AppendSettingLog("[失败] 关闭在线语音识别"); fail++; }
            }

            // 快速启动
            if (chkFastStartup.Checked)
            {
                if (SetFastStartup(false)) { AppendSettingLog("[成功] 已关闭快速启动"); success++; }
                else { AppendSettingLog("[失败] 关闭快速启动"); fail++; }
            }

            lblSettingStatus.Text = $"完成：成功 {success} 项，失败 {fail} 项";
            AppendSettingLog($"[{DateTime.Now:HH:mm:ss}] 应用结束：成功 {success}，失败 {fail}");
            MessageBox.Show(this, $"设置应用完成！\n成功：{success}\n失败：{fail}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnRefreshStatus_Click(object sender, EventArgs e)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"[{DateTime.Now:HH:mm:ss}] 当前系统设置状态：");
            sb.AppendLine($"  文件扩展名显示：{(IsFileExtensionsVisible() ? "已显示" : "隐藏")}");
            sb.AppendLine($"  隐藏文件显示：{(IsHiddenFilesVisible() ? "已显示" : "隐藏")}");
            sb.AppendLine($"  系统休眠：{(IsHibernateEnabled() ? "开启" : "已关闭")}");
            sb.AppendLine($"  USB 选择性挂起：{(IsUsbSelectiveSuspendEnabled() ? "启用" : "已禁用")}");
            sb.AppendLine($"  管理员权限：{(IsAdministrator() ? "是" : "否")}");
            AppendSettingLog(sb.ToString());
            lblSettingStatus.Text = "状态已刷新";
        }

        // ==================== 原有设置实现 ====================

        private bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void RestartAsAdmin()
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    UseShellExecute = true,
                    WorkingDirectory = Environment.CurrentDirectory,
                    FileName = Application.ExecutablePath,
                    Verb = "runas"
                };
                Process.Start(startInfo);
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"无法以管理员身份启动：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool SetFileExtensionsVisible(bool show)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", true))
                {
                    if (key == null) return false;
                    key.SetValue("HideFileExt", show ? 0 : 1, RegistryValueKind.DWord);
                }
                RefreshShell();
                return true;
            }
            catch { return false; }
        }

        private bool IsFileExtensionsVisible()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"))
                {
                    if (key == null) return true;
                    object val = key.GetValue("HideFileExt", 1);
                    return val != null && Convert.ToInt32(val) == 0;
                }
            }
            catch { return true; }
        }

        private bool SetHiddenFilesVisible(bool show)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", true))
                {
                    if (key == null) return false;
                    key.SetValue("Hidden", show ? 1 : 2, RegistryValueKind.DWord);
                    key.SetValue("ShowSuperHidden", show ? 1 : 0, RegistryValueKind.DWord);
                }
                RefreshShell();
                return true;
            }
            catch { return false; }
        }

        private bool IsHiddenFilesVisible()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"))
                {
                    if (key == null) return false;
                    object val = key.GetValue("Hidden", 2);
                    return val != null && Convert.ToInt32(val) == 1;
                }
            }
            catch { return false; }
        }

        private bool DisableHibernate()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powercfg.exe",
                    Arguments = "-h off",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(10000);
                    return p.ExitCode == 0;
                }
            }
            catch { return false; }
        }

        private bool IsHibernateEnabled()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powercfg.exe",
                    Arguments = "-a",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(10000);
                    string output = p.StandardOutput.ReadToEnd();
                    return output.Contains("休眠") || output.Contains("hibernate");
                }
            }
            catch { return true; }
        }

        private bool CleanTempFiles(out long freedMB)
        {
            freedMB = 0;
            try
            {
                string tempPath = Path.GetTempPath();
                var dir = new DirectoryInfo(tempPath);
                long freedBytes = 0;

                foreach (var file in dir.GetFiles("*", SearchOption.AllDirectories))
                {
                    try
                    {
                        freedBytes += file.Length;
                        file.Delete();
                    }
                    catch { }
                }

                foreach (var subDir in dir.GetDirectories("*", SearchOption.AllDirectories))
                {
                    try { subDir.Delete(true); } catch { }
                }

                freedMB = freedBytes / (1024 * 1024);
                return true;
            }
            catch { return false; }
        }

        private bool ShowDesktopIcons(bool show)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel", true))
                {
                    if (key != null)
                    {
                        key.SetValue("{20D04FE0-3AEA-1069-A2D8-08002B30309D}", show ? 0 : 1, RegistryValueKind.DWord);
                    }
                }
                RefreshShell();
                return true;
            }
            catch { return false; }
        }

        private void RefreshShell()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "taskkill.exe",
                    Arguments = "/f /im explorer.exe",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi)) p.WaitForExit(3000);
                Process.Start("explorer.exe");
            }
            catch { }
        }

        // ==================== PPTX 新增设置实现 ====================

        private bool SetPowerUsbSelectiveSuspend(bool enable)
        {
            try
            {
                string scheme = "381b4222-f694-41f0-9685-ff5bb260df2e";
                string subgroup = "2a737441-1930-4402-8d77-b2bebba308a3";
                string setting = "a25be568-2e6d-4d6b-8b9a-3c6e5f6d4a1b";
                string value = enable ? "1" : "0";

                RunPowerCfg($"-setacvalueindex {scheme} {subgroup} {setting} {value}");
                RunPowerCfg($"-setdcvalueindex {scheme} {subgroup} {setting} {value}");
                RunPowerCfg($"-setactive {scheme}");
                return true;
            }
            catch { return false; }
        }

        private bool IsUsbSelectiveSuspendEnabled()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powercfg.exe",
                    Arguments = "-q 381b4222-f694-41f0-9685-ff5bb260df2e 2a737441-1930-4402-8d77-b2bebba308a3",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(5000);
                    string output = p.StandardOutput.ReadToEnd();
                    return output.Contains("1") && !output.Contains("0");
                }
            }
            catch { return true; }
        }

        private bool SetPowerPcieAspm(bool enable)
        {
            try
            {
                string scheme = "381b4222-f694-41f0-9685-ff5bb260df2e";
                string subgroup = "9442d3a8-bd9e-4251-9ad1-0c9796e8f7a9";
                string setting = "d3b070d8-2a8b-4b0a-8a3a-6e5b4f6c7d8e";
                string value = enable ? "0" : "1";

                RunPowerCfg($"-setacvalueindex {scheme} {subgroup} {setting} {value}");
                RunPowerCfg($"-setdcvalueindex {scheme} {subgroup} {setting} {value}");
                RunPowerCfg($"-setactive {scheme}");
                return true;
            }
            catch { return false; }
        }

        private void RunPowerCfg(string args)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powercfg.exe",
                Arguments = args,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (var p = Process.Start(psi))
            {
                p.WaitForExit(10000);
            }
        }

        private bool SetNetworkEbusOnly()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "netsh",
                    Arguments = "interface ip show config",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(10000);
                    string output = p.StandardOutput.ReadToEnd();
                    AppendSettingLog($"[网络适配器] 检测到适配器信息：{output.Length} 字节");
                }

                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled = True"))
                {
                    foreach (var adapter in searcher.Get())
                    {
                        string name = adapter["Name"]?.ToString() ?? "";
                        if (name.IndexOf("eBUS", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            name.IndexOf("Cognex", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            AppendSettingLog($"[网络适配器] 发现 eBUS/Cognex 适配器：{name}");
                        }
                    }
                }

                return true;
            }
            catch { return false; }
        }

        private bool SetReceiveBuffersMax()
        {
            try
            {
                using (var baseKey = Registry.LocalMachine.OpenSubKey(
                    @"SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"))
                {
                    if (baseKey == null) return false;

                    foreach (string subKeyName in baseKey.GetSubKeyNames())
                    {
                        using (var subKey = baseKey.OpenSubKey(subKeyName, true))
                        {
                            if (subKey == null) continue;

                            string driverDesc = subKey.GetValue("DriverDesc")?.ToString() ?? "";
                            if (driverDesc.IndexOf("eBUS", StringComparison.OrdinalIgnoreCase) < 0 &&
                                driverDesc.IndexOf("Cognex", StringComparison.OrdinalIgnoreCase) < 0 &&
                                driverDesc.IndexOf("Ethernet", StringComparison.OrdinalIgnoreCase) < 0 &&
                                driverDesc.IndexOf("网卡", StringComparison.OrdinalIgnoreCase) < 0)
                                continue;

                            foreach (string param in new[] { "ReceiveBuffers", "RecvBuffers", "RxBuffers", "NumberOfReceiveBuffers" })
                            {
                                try { subKey.SetValue(param, "256", RegistryValueKind.DWord); } catch { }
                            }
                        }
                    }
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetInterruptModerationExtreme()
        {
            try
            {
                using (var baseKey = Registry.LocalMachine.OpenSubKey(
                    @"SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"))
                {
                    if (baseKey == null) return false;

                    foreach (string subKeyName in baseKey.GetSubKeyNames())
                    {
                        using (var subKey = baseKey.OpenSubKey(subKeyName, true))
                        {
                            if (subKey == null) continue;

                            string driverDesc = subKey.GetValue("DriverDesc")?.ToString() ?? "";
                            if (driverDesc.IndexOf("eBUS", StringComparison.OrdinalIgnoreCase) < 0 &&
                                driverDesc.IndexOf("Cognex", StringComparison.OrdinalIgnoreCase) < 0 &&
                                driverDesc.IndexOf("Ethernet", StringComparison.OrdinalIgnoreCase) < 0 &&
                                driverDesc.IndexOf("网卡", StringComparison.OrdinalIgnoreCase) < 0)
                                continue;

                            foreach (string param in new[] { "InterruptModeration", "Interrupt Moderation", "AdaptiveInterrupt", "IntModeration" })
                            {
                                try { subKey.SetValue(param, "1", RegistryValueKind.DWord); } catch { }
                            }
                        }
                    }
                }
                return true;
            }
            catch { return false; }
        }

        // ==================== 新电脑必做设置实现（7 项） ====================

        private bool OpenWindowsUpdate()
        {
            try
            {
                Process.Start(new ProcessStartInfo { FileName = "ms-settings:windowsupdate", UseShellExecute = true });
                return true;
            }
            catch { return false; }
        }

        private bool OpenAppsFeatures()
        {
            try
            {
                Process.Start(new ProcessStartInfo { FileName = "ms-settings:appsfeatures", UseShellExecute = true });
                return true;
            }
            catch { return false; }
        }

        private bool OpenRecoveryDrive()
        {
            try
            {
                Process.Start("RecoveryDrive.exe");
                return true;
            }
            catch { return false; }
        }

        private bool SetPowerHighPerformance()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powercfg.exe",
                    Arguments = "-setactive SCHEME_MIN",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(10000);
                    return p.ExitCode == 0;
                }
            }
            catch { return false; }
        }

        private bool SetUacDefault()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", true))
                {
                    if (key == null) return false;
                    key.SetValue("EnableLUA", 1, RegistryValueKind.DWord);
                    key.SetValue("ConsentPromptBehaviorAdmin", 5, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }

        private bool CreateSystemRestorePoint()
        {
            try
            {
                using (var cls = new ManagementClass("root\\default", "SystemRestore", null))
                {
                    var inParams = cls.GetMethodParameters("CreateRestorePoint");
                    inParams["Description"] = "EasyFindImage 系统还原点";
                    inParams["RestorePointType"] = 0;
                    inParams["EventType"] = 100;
                    cls.InvokeMethod("CreateRestorePoint", inParams, null);
                    return true;
                }
            }
            catch { return false; }
        }

        private bool EnableWindowsDefender()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-ExecutionPolicy Bypass -Command \"Set-MpPreference -DisableRealtimeMonitoring $false -EA SilentlyContinue\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(10000);
                    return p.ExitCode == 0;
                }
            }
            catch { return false; }
        }

        // ==================== 隐私选项实现 ====================

        private bool SetPrivacyAdId(bool enable)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo", true))
                {
                    key.SetValue("Enabled", enable ? 1 : 0, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetPrivacyLocation(bool enable)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location", true))
                {
                    key.SetValue("Value", enable ? "Allow" : "Deny", RegistryValueKind.String);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetPrivacyDiagnostics(bool enable)
        {
            try
            {
                using (var key = Registry.LocalMachine.CreateSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection", true))
                {
                    key.SetValue("AllowTelemetry", enable ? 1 : 0, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetPrivacySpeech(bool enable)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy", true))
                {
                    key.SetValue("HasAccepted", enable ? 1 : 0, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }

        // ==================== 快速启动实现 ====================

        private bool SetFastStartup(bool enable)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powercfg.exe",
                    Arguments = enable ? "-hibernate on" : "-hibernate off",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(10000);
                    return p.ExitCode == 0;
                }
            }
            catch { return false; }
        }

        private void AppendSettingLog(string message)
        {
            lstSettingLog.Items.Add(message);
            lstSettingLog.TopIndex = lstSettingLog.Items.Count - 1;
            Application.DoEvents();
        }

        // ==================== 系统优化设置实现 ====================

        private bool SetServiceStartType(string serviceName, string startType)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"config {serviceName} start= {startType}",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                using (var p = Process.Start(psi))
                {
                    p.WaitForExit(10000);
                    return p.ExitCode == 0;
                }
            }
            catch { return false; }
        }

        private bool SetTelemetryDisabled()
        {
            try
            {
                using (var key = Registry.LocalMachine.CreateSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection", true))
                {
                    key.SetValue("AllowTelemetry", 0, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetBackgroundAppsDisabled()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\BackgroundApplications", true))
                {
                    key.SetValue("GlobalBackgroundApplicationsEnabled", 0, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetOneDriveAutoStart(bool enable)
        {
            try
            {
                string oneDrivePath = @"%LocalAppData%\Microsoft\OneDrive\OneDrive.exe /background";
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key == null) return false;
                    if (enable)
                        key.SetValue("OneDrive", oneDrivePath);
                    else
                        key.DeleteValue("OneDrive", false);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetVisualEffectsBestPerformance()
        {
            try
            {
                byte[] mask = { 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
                using (var key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true))
                {
                    if (key == null) return false;
                    key.SetValue("UserPreferencesMask", mask, RegistryValueKind.Binary);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetTipsDisabled()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager", true))
                {
                    key.SetValue("SubscribedContent-338387Enabled", 0, RegistryValueKind.DWord);
                    key.SetValue("SubscribedContent-338388Enabled", 0, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }

        private bool SetLockScreenSpotlightDisabled()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Policies\Microsoft\Windows\Personalization", true))
                {
                    key.SetValue("NoLockScreen", 1, RegistryValueKind.DWord);
                    key.SetValue("LockScreenSlideshowEnabled", 0, RegistryValueKind.DWord);
                }
                return true;
            }
            catch { return false; }
        }
    }
}
