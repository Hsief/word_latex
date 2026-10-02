using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

namespace WordLatexRepair
{
    /// <summary>Repairs only this add-in's registration for the user who launches the installer.</summary>
    internal static class Program
    {
        private const string AddinKey = @"Software\Microsoft\Office\Word\Addins\WordLatexVSTOAddin";
        private static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordLatexVSTO", "diagnostics");

        [STAThread]
        private static int Main(string[] args)
        {
            bool quiet = args.Any(a => string.Equals(a, "/quiet", StringComparison.OrdinalIgnoreCase));
            bool diagnose = args.Any(a => string.Equals(a, "/diagnose", StringComparison.OrdinalIgnoreCase));
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                if (identity.IsSystem || identity.User == null)
                    throw new InvalidOperationException("请在实际使用 Word 的 Windows 用户下运行，不要使用系统服务账户。");

                string manifest = FindInstalledManifest();
                string uri = new Uri(manifest).AbsoluteUri + "|vstolocal";
                Directory.CreateDirectory(LogDirectory);
                string account = identity.Name + " (" + identity.User.Value + ")";
                if (!diagnose)
                {
                    int session = Process.GetCurrentProcess().SessionId;
                    bool wordRunning = Process.GetProcessesByName("WINWORD").Any(p => p.SessionId == session);
                    if (wordRunning)
                        throw new InvalidOperationException("请先保存文档并关闭所有 Word 窗口，再运行修复。不会强制关闭 Word。");

                    BackupRegistration(account);
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(AddinKey))
                    {
                        if (key == null) throw new IOException("无法写入当前用户的 Word 加载项注册项。");
                        key.SetValue("FriendlyName", "WordLatexVSTO", RegistryValueKind.String);
                        key.SetValue("Description", "LaTeX 与 Word 原生公式双向转换", RegistryValueKind.String);
                        key.SetValue("Manifest", uri, RegistryValueKind.String);
                        key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    }
                }

                string current;
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(AddinKey))
                {
                    current = key == null ? "HKCU: 未设置（使用系统注册项）" :
                        "HKCU LoadBehavior=" + key.GetValue("LoadBehavior", "未设置") +
                        Environment.NewLine + "HKCU Manifest=" + key.GetValue("Manifest", "未设置");
                    if (!diagnose && (key == null || !Equals(key.GetValue("LoadBehavior"), 3) ||
                        !Equals(key.GetValue("Manifest"), uri)))
                        throw new IOException("当前用户注册项写入后校验失败。");
                }
                string message = (diagnose ? "注册诊断完成。" : "当前用户自动加载注册已修复。请重新打开 Word。") +
                    Environment.NewLine + "账户：" + account + Environment.NewLine +
                    "安装清单：" + manifest + Environment.NewLine + current + Environment.NewLine +
                    "这不是实际 Ribbon 加载验证；请确认 Word 顶部出现“论文工具”。";
                File.WriteAllText(Path.Combine(LogDirectory, "registration-latest.txt"),
                    DateTimeOffset.Now.ToString("o") + Environment.NewLine + message);
                if (!quiet) MessageBox.Show(message, "WordLatexVSTO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            catch (Exception exception)
            {
                try
                {
                    Directory.CreateDirectory(LogDirectory);
                    File.WriteAllText(Path.Combine(LogDirectory, "registration-error.txt"),
                        DateTimeOffset.Now.ToString("o") + Environment.NewLine + exception.ToString());
                }
                catch { /* A logging failure must not conceal the original registration error. */ }
                if (!quiet) MessageBox.Show(exception.Message, "WordLatexVSTO 修复失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        /// <summary>Uses the MSI's machine registration, never an obsolete per-user development path.</summary>
        private static string FindInstalledManifest()
        {
            foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (RegistryKey machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (RegistryKey key = machine.OpenSubKey(AddinKey))
                {
                    string value = key == null ? null : key.GetValue("Manifest") as string;
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    const string suffix = "|vstolocal";
                    if (!value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) continue;
                    Uri uri;
                    if (!Uri.TryCreate(value.Substring(0, value.Length - suffix.Length), UriKind.Absolute, out uri) || !uri.IsFile)
                        continue;
                    string manifest = Path.GetFullPath(uri.LocalPath);
                    string directory = Path.GetDirectoryName(manifest);
                    if (!string.Equals(Path.GetFileName(manifest), "WordLatexVSTOAddin.vsto", StringComparison.OrdinalIgnoreCase)) continue;
                    if (new[] { manifest, Path.Combine(directory, "WordLatexVSTOAddin.dll"),
                        Path.Combine(directory, "WordLatexVSTOAddin.dll.manifest"), Path.Combine(directory, "AfterMathCore.dll") }.All(File.Exists))
                        return manifest;
                }
            }
            throw new FileNotFoundException("未找到完整的 MSI 安装。请先运行 WordLatexVSTO_Setup.exe，而不是单独打开 .vsto 文件。");
        }

        /// <summary>Saves the previous values of this single add-in before changing them; other add-ins stay untouched.</summary>
        private static void BackupRegistration(string account)
        {
            XElement snapshot = new XElement("RegistrationBackup", new XAttribute("account", account),
                new XAttribute("key", "HKEY_CURRENT_USER\\" + AddinKey));
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(AddinKey))
            {
                snapshot.SetAttributeValue("existed", key != null);
                if (key != null)
                {
                    foreach (string name in new[] { "FriendlyName", "Description", "Manifest", "LoadBehavior" })
                    {
                        object value = key.GetValue(name);
                        if (value != null) snapshot.Add(new XElement("Value", new XAttribute("name", name),
                            new XAttribute("kind", key.GetValueKind(name)), value.ToString()));
                    }
                }
            }
            snapshot.Save(Path.Combine(LogDirectory, "registration-backup-" +
                DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff") + ".xml"));
        }
    }
}
