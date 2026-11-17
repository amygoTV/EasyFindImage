using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace EasyFindImage
{
    /// <summary>
    /// 程序配置持久化类，使用 XML 文件保存用户设置
    /// </summary>
    [Serializable]
    public class AppConfig
    {
        // ==================== Tab 1：表格汇总与筛选（仅保存开关状态，不保存路径） ====================
        public string FilterColumn { get; set; } = "";
        public string FilterValue { get; set; } = "";
        public List<string> CleanConditions { get; set; } = new List<string>();
        public bool RemoveDuplicateSN { get; set; } = false;
        public bool KeepOnlyCommonCols { get; set; } = false;

        // ==================== Tab 2：按SN查找图片（仅保存开关状态，不保存路径） ====================
        public bool SnFromExcel { get; set; } = true;
        public bool SearchAllDrives { get; set; } = false;

        // ==================== Tab 3：一键电脑设置 ====================
        // 原有设置
        public bool ShowExt { get; set; } = true;
        public bool ShowHidden { get; set; } = false;
        public bool HibernateOff { get; set; } = false;
        public bool CleanTemp { get; set; } = true;
        public bool DesktopIcons { get; set; } = false;

        // PPTX 新增设置
        public bool UsbPowerDisabled { get; set; } = false;
        public bool PciePowerDisabled { get; set; } = false;
        public bool NetworkEbusOnly { get; set; } = false;
        public bool ReceiveBuffersMax { get; set; } = false;
        public bool InterruptModerationExtreme { get; set; } = false;
        public bool BiosPcieAutoReminder { get; set; } = false;

        // 新电脑必做设置
        public bool WinUpdate { get; set; } = false;
        public bool RestorePoint { get; set; } = false;
        public bool Defender { get; set; } = false;
        public bool RemoveBloatware { get; set; } = false;
        public bool PowerHighPerf { get; set; } = false;
        public bool UAC { get; set; } = false;
        public bool RecoveryDrive { get; set; } = false;

        // 系统优化
        public bool OptSysMain { get; set; } = false;
        public bool OptWindowsSearch { get; set; } = false;
        public bool OptTelemetry { get; set; } = false;
        public bool OptBackgroundApps { get; set; } = false;
        public bool OptOneDrive { get; set; } = false;
        public bool OptVisualEffects { get; set; } = false;
        public bool OptTips { get; set; } = false;
        public bool OptLockScreen { get; set; } = false;

        // 隐私选项
        public bool PrivAdId { get; set; } = false;
        public bool PrivLocation { get; set; } = false;
        public bool PrivDiagnostics { get; set; } = false;
        public bool PrivSpeech { get; set; } = false;

        // 快速启动
        public bool FastStartup { get; set; } = false;
    }

    public static class ConfigHelper
    {
        private static readonly string ConfigFilePath;

        static ConfigHelper()
        {
            // 配置文件保存在程序启动目录（exe 所在目录）
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            ConfigFilePath = Path.Combine(exeDir, "config.xml");
        }

        public static string ConfigPath => ConfigFilePath;

        /// <summary>保存配置到本地 XML 文件</summary>
        public static void Save(AppConfig config)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(AppConfig));
                using (var writer = new StreamWriter(ConfigFilePath, false, System.Text.Encoding.UTF8))
                {
                    serializer.Serialize(writer, config);
                }
            }
            catch (Exception ex)
            {
                // 保存失败不阻塞程序，仅输出到日志（可选）
                System.Diagnostics.Debug.WriteLine($"配置保存失败：{ex.Message}");
            }
        }

        /// <summary>从本地 XML 文件加载配置</summary>
        public static AppConfig Load()
        {
            try
            {
                if (!File.Exists(ConfigFilePath))
                    return new AppConfig();

                var serializer = new XmlSerializer(typeof(AppConfig));
                using (var reader = new StreamReader(ConfigFilePath, System.Text.Encoding.UTF8))
                {
                    var config = serializer.Deserialize(reader) as AppConfig;
                    if (config != null)
                        return config;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"配置加载失败：{ex.Message}");
            }
            return new AppConfig();
        }
    }
}
