using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace EMergePortable
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.ToString(),
                    "eMerge Portable",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static void Run()
        {
            string workingDirectory = GetWorkingDirectory();
            string extractedRoot = EnsurePayloadExtracted();
            string appPath = Path.Combine(extractedRoot, "eMerge.exe");
            if (!File.Exists(appPath))
            {
                throw new FileNotFoundException("解包后未找到 eMerge.exe。", appPath);
            }

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = appPath;
            startInfo.Arguments = BuildArgumentString(Environment.GetCommandLineArgs(), 1);
            startInfo.WorkingDirectory = Directory.Exists(workingDirectory) ? workingDirectory : extractedRoot;
            startInfo.UseShellExecute = false;

            Process child = Process.Start(startInfo);
            if (child == null)
            {
                throw new InvalidOperationException("启动 eMerge 失败。");
            }
        }

        private static string GetWorkingDirectory()
        {
            try
            {
                return Environment.CurrentDirectory;
            }
            catch
            {
                return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
            }
        }

        private static string EnsurePayloadExtracted()
        {
            string launcherPath = Assembly.GetExecutingAssembly().Location;
            FileInfo launcherInfo = new FileInfo(launcherPath);
            string versionToken = string.Format(
                "payload-{0}-{1}",
                launcherInfo.Length,
                launcherInfo.LastWriteTimeUtc.Ticks);
            string rootDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "eMergePortable");
            string targetDir = Path.Combine(rootDir, versionToken);
            string readyMarker = Path.Combine(targetDir, ".ready");

            if (File.Exists(readyMarker) && File.Exists(Path.Combine(targetDir, "eMerge.exe")))
            {
                return targetDir;
            }

            Directory.CreateDirectory(rootDir);

            string tempDir = targetDir + ".tmp";
            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, true);
            }
            Directory.CreateDirectory(tempDir);

            string payloadZipPath = Path.Combine(tempDir, "payload.zip");
            using (Stream payloadStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip"))
            {
                if (payloadStream == null)
                {
                    throw new InvalidOperationException("未找到嵌入的 payload.zip 资源。");
                }
                using (FileStream output = File.Create(payloadZipPath))
                {
                    payloadStream.CopyTo(output);
                }
            }

            ZipFile.ExtractToDirectory(payloadZipPath, tempDir);
            File.Delete(payloadZipPath);

            if (Directory.Exists(targetDir))
            {
                Directory.Delete(targetDir, true);
            }
            Directory.Move(tempDir, targetDir);
            File.WriteAllText(readyMarker, "ok", Encoding.UTF8);
            return targetDir;
        }

        private static string BuildArgumentString(string[] args, int skipCount)
        {
            if (args == null || args.Length <= skipCount)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            for (int i = skipCount; i < args.Length; i += 1)
            {
                if (builder.Length > 0)
                {
                    builder.Append(' ');
                }
                builder.Append(QuoteArgument(args[i]));
            }
            return builder.ToString();
        }

        private static string QuoteArgument(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "\"\"";
            }

            bool needsQuotes = false;
            for (int i = 0; i < value.Length; i += 1)
            {
                char ch = value[i];
                if (char.IsWhiteSpace(ch) || ch == '"')
                {
                    needsQuotes = true;
                    break;
                }
            }
            if (!needsQuotes)
            {
                return value;
            }

            StringBuilder builder = new StringBuilder();
            builder.Append('"');
            int backslashCount = 0;
            for (int i = 0; i < value.Length; i += 1)
            {
                char ch = value[i];
                if (ch == '\\')
                {
                    backslashCount += 1;
                    continue;
                }

                if (ch == '"')
                {
                    builder.Append('\\', backslashCount * 2 + 1);
                    builder.Append('"');
                    backslashCount = 0;
                    continue;
                }

                if (backslashCount > 0)
                {
                    builder.Append('\\', backslashCount);
                    backslashCount = 0;
                }
                builder.Append(ch);
            }

            if (backslashCount > 0)
            {
                builder.Append('\\', backslashCount * 2);
            }
            builder.Append('"');
            return builder.ToString();
        }
    }
}
