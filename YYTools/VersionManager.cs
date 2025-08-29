using System;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.IO;

namespace YYTools
{
    /// <summary>
    /// 版本管理器
    /// 统一管理版本号、构建信息和哈希值
    /// </summary>
    public static class VersionManager
    {
        #region 版本信息常量

        /// <summary>
        /// 主版本号
        /// </summary>
        public const int MajorVersion = 3;

        /// <summary>
        /// 次版本号
        /// </summary>
        public const int MinorVersion = 0;

        /// <summary>
        /// 修订版本号
        /// </summary>
        public const int PatchVersion = 0;

        /// <summary>
        /// 构建版本号（自动递增）
        /// </summary>
        public static readonly int BuildVersion = GetBuildNumber();

        /// <summary>
        /// 版本标识（Release/Beta/Alpha）
        /// </summary>
        public const string VersionTag = "Release";

        /// <summary>
        /// 构建日期
        /// </summary>
        public static readonly DateTime BuildDate = GetBuildDate();

        #endregion

        #region 版本属性

        /// <summary>
        /// 完整版本号字符串
        /// </summary>
        public static string FullVersion => $"{MajorVersion}.{MinorVersion}.{PatchVersion}.{BuildVersion}";

        /// <summary>
        /// 显示版本号（不包含构建号）
        /// </summary>
        public static string DisplayVersion => $"{MajorVersion}.{MinorVersion}.{PatchVersion}";

        /// <summary>
        /// 版本号带标识
        /// </summary>
        public static string VersionWithTag => string.IsNullOrEmpty(VersionTag) ? DisplayVersion : $"{DisplayVersion}-{VersionTag}";

        /// <summary>
        /// 程序集版本
        /// </summary>
        public static Version AssemblyVersion => Assembly.GetExecutingAssembly().GetName().Version;

        /// <summary>
        /// 文件版本
        /// </summary>
        public static string FileVersion => GetFileVersion();

        #endregion

        #region 哈希值和唯一标识

        /// <summary>
        /// 程序集哈希值（SHA256）
        /// </summary>
        public static string AssemblyHash => GetAssemblyHash();

        /// <summary>
        /// 版本唯一标识符
        /// </summary>
        public static string VersionGuid => GetVersionGuid();

        /// <summary>
        /// 构建唯一标识符
        /// </summary>
        public static string BuildGuid => GetBuildGuid();

        #endregion

        #region 版本信息方法

        /// <summary>
        /// 获取完整版本信息
        /// </summary>
        public static string GetFullVersionInfo()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"产品名称: YY运单匹配工具");
            sb.AppendLine($"版本号: {VersionWithTag}");
            sb.AppendLine($"完整版本: {FullVersion}");
            sb.AppendLine($"构建日期: {BuildDate:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"程序集版本: {AssemblyVersion}");
            sb.AppendLine($"文件版本: {FileVersion}");
            sb.AppendLine($"版本GUID: {VersionGuid}");
            sb.AppendLine($"构建GUID: {BuildGuid}");
            sb.AppendLine($"程序集哈希: {AssemblyHash}");
            sb.AppendLine($"框架版本: {Environment.Version}");
            sb.AppendLine($"操作系统: {Environment.OSVersion}");
            sb.AppendLine($"处理器架构: {Environment.Is64BitProcess}位");
            return sb.ToString();
        }

        /// <summary>
        /// 获取简短版本信息
        /// </summary>
        public static string GetShortVersionInfo()
        {
            return $"YY运单匹配工具 v{VersionWithTag} (Build {BuildVersion})";
        }

        /// <summary>
        /// 获取版本比较信息
        /// </summary>
        public static VersionInfo GetVersionInfo()
        {
            return new VersionInfo
            {
                Major = MajorVersion,
                Minor = MinorVersion,
                Patch = PatchVersion,
                Build = BuildVersion,
                Tag = VersionTag,
                BuildDate = BuildDate,
                FullVersion = FullVersion,
                DisplayVersion = DisplayVersion,
                VersionGuid = VersionGuid,
                BuildGuid = BuildGuid,
                AssemblyHash = AssemblyHash
            };
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 获取构建号
        /// </summary>
        private static int GetBuildNumber()
        {
            try
            {
                // 基于当前日期生成构建号，确保递增
                var baseDate = new DateTime(2025, 1, 1);
                var daysSince = (DateTime.Now - baseDate).Days;
                var minutesToday = DateTime.Now.Hour * 60 + DateTime.Now.Minute;
                return daysSince * 1000 + minutesToday / 10; // 每10分钟递增1
            }
            catch
            {
                return 1;
            }
        }

        /// <summary>
        /// 获取构建日期
        /// </summary>
        private static DateTime GetBuildDate()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var fileInfo = new FileInfo(assembly.Location);
                return fileInfo.LastWriteTime;
            }
            catch
            {
                return DateTime.Now;
            }
        }

        /// <summary>
        /// 获取文件版本
        /// </summary>
        private static string GetFileVersion()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var fileVersionAttribute = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>();
                return fileVersionAttribute?.Version ?? FullVersion;
            }
            catch
            {
                return FullVersion;
            }
        }

        /// <summary>
        /// 获取程序集哈希值
        /// </summary>
        private static string GetAssemblyHash()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var assemblyPath = assembly.Location;
                
                if (!File.Exists(assemblyPath))
                    return "未知";

                using (var sha256 = SHA256.Create())
                {
                    var fileBytes = File.ReadAllBytes(assemblyPath);
                    var hashBytes = sha256.ComputeHash(fileBytes);
                    return BitConverter.ToString(hashBytes).Replace("-", "").Substring(0, 16); // 取前16位
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"获取程序集哈希失败: {ex.Message}");
                return "计算失败";
            }
        }

        /// <summary>
        /// 获取版本GUID
        /// </summary>
        private static string GetVersionGuid()
        {
            try
            {
                // 基于版本信息生成确定性GUID
                var versionString = $"{MajorVersion}.{MinorVersion}.{PatchVersion}-{VersionTag}";
                using (var md5 = MD5.Create())
                {
                    var hashBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(versionString));
                    var guid = new Guid(hashBytes);
                    return guid.ToString("D").ToUpper();
                }
            }
            catch
            {
                return Guid.NewGuid().ToString("D").ToUpper();
            }
        }

        /// <summary>
        /// 获取构建GUID
        /// </summary>
        private static string GetBuildGuid()
        {
            try
            {
                // 基于完整版本和构建日期生成确定性GUID
                var buildString = $"{FullVersion}-{BuildDate:yyyyMMddHHmm}";
                using (var md5 = MD5.Create())
                {
                    var hashBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(buildString));
                    var guid = new Guid(hashBytes);
                    return guid.ToString("D").ToUpper();
                }
            }
            catch
            {
                return Guid.NewGuid().ToString("D").ToUpper();
            }
        }

        #endregion

        #region 版本比较

        /// <summary>
        /// 比较版本号
        /// </summary>
        public static int CompareVersion(string version1, string version2)
        {
            try
            {
                var v1 = new Version(version1);
                var v2 = new Version(version2);
                return v1.CompareTo(v2);
            }
            catch
            {
                return string.Compare(version1, version2, StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// 检查是否为新版本
        /// </summary>
        public static bool IsNewerVersion(string otherVersion)
        {
            return CompareVersion(FullVersion, otherVersion) > 0;
        }

        #endregion
    }

    /// <summary>
    /// 版本信息结构体
    /// </summary>
    public class VersionInfo
    {
        public int Major { get; set; }
        public int Minor { get; set; }
        public int Patch { get; set; }
        public int Build { get; set; }
        public string Tag { get; set; }
        public DateTime BuildDate { get; set; }
        public string FullVersion { get; set; }
        public string DisplayVersion { get; set; }
        public string VersionGuid { get; set; }
        public string BuildGuid { get; set; }
        public string AssemblyHash { get; set; }

        public override string ToString()
        {
            return $"{DisplayVersion}-{Tag} (Build {Build})";
        }
    }
}
