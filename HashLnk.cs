using IWshRuntimeLibrary;
using Microsoft.WindowsAPICodePack.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;



namespace MenuWinX
{
    public class HashLnk
    {
        /*
         * 作者：坑晨
         * 适用范围：Windows8 +
         * Website：http://www.pcmoe.net/forum.php
         * 
         * 作者：riverar
         * WinX Menu Hashing Algorithm：https://www.withinrafael.com/2014/04/05/the-winx-menu-and-its-hashing-algorithm/
         * PropertyStoreDataBlock：https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-shllink/36463387-0708-40f6-a3a5-452fe42be585
         */

        /// <summary>
        /// 快速添加 Hashing
        /// </summary>
        public HashLnk(string lnkFile)
        {
            PropertyStoreDataBlock(lnkFile);
        }

        /// <summary>
        /// 快速批量添加 Hashing
        /// </summary>
        public HashLnk(string[] lnkFiles)
        {
            foreach (string lnk in lnkFiles)
            {
                PropertyStoreDataBlock(lnk);
            }
        }

        /// <summary>
        /// 快速批量添加 Hashing
        /// </summary>
        public HashLnk(List<string> lnkFiles)
        {
            foreach (string lnk in lnkFiles)
            {
                PropertyStoreDataBlock(lnk);
            }
        }

        /// <summary>
        /// 创建快捷方式并添加 Hashing
        /// </summary>
        public static void CreateLnk(int group, string displayName, string exePath, string workingDir, bool elevated, string arguments)
        {
            _groupFolderPath = Path.Combine(WinXFolder, "Group" + group);
            Directory.CreateDirectory(_groupFolderPath);

            string lnkPath = Path.Combine(_groupFolderPath, displayName + ".lnk");

            IWshShortcut wshShortcut = (IWshShortcut)wshShell.CreateShortcut(lnkPath);
            wshShortcut.Description = displayName;
            wshShortcut.TargetPath = ExpandEnvVar(exePath);
            wshShortcut.WorkingDirectory = workingDir;
            wshShortcut.Arguments = ExpandEnvVar(arguments);
            wshShortcut.Save();

            if (elevated) { ElevateLnk(lnkPath); }

            PropertyStoreDataBlock(lnkPath);
        }

        /// <summary>
        /// 检查快捷方式文件格式
        /// </summary>
        public static bool IsNoLnkFile(string lnkFile)
        {
            if (Path.GetExtension(lnkFile)?.ToLower() != ".lnk") { return true; }
            else { return false; }
        }

        /// <summary>
        /// 添加 Hashing
        /// </summary>
        public static bool PropertyStoreDataBlock(string lnkFile)
        {
            IWshShortcut lnk = (IWshShortcut)wshShell.CreateShortcut(lnkFile);
            string property = lnk.TargetPath;

            foreach (var kv in SystemFolderMapping)
            {
                //执行快捷方式映射操作
                property = property.ToLower().Replace(kv.Key, kv.Value);
            }

            if (lnk.Arguments.Length > 0) { property += lnk.Arguments; }
            property += "do not prehash links.  this should only be done by the user."; //特殊但必须存在的字符串
            property = property.ToLower();

            byte[] inBytes = Encoding.GetEncoding(1200).GetBytes(property);
            int byteCount = inBytes.Length;
            byte[] outBytes = new byte[byteCount];

            int hashResult = HashData(inBytes, byteCount, outBytes, byteCount);
            if (hashResult != 0) { return false; } //输出错误 Marshal.GetLastWin32Error()

            using (var propertyWriter = ShellFile.FromFilePath(lnkFile).Properties.GetPropertyWriter())
            {
                //GUID {7B2D8DFB-D190-344E-BF60-6EAC09922BBF}
                //DEFINE_PROPERTYKEY(PKEY_WINX_HASH, 0xFB8D2D7B, 0x90D1, 0x4E34, 0xBF, 0x60, 0x6E, 0xAC, 0x09, 0x92, 0x2B, 0xBF, 0x02);
                propertyWriter.WriteProperty("System.Winx.Hash", BitConverter.ToUInt32(outBytes, 0));
            }

            return true;
        }

        /// <summary>
        /// 获取 Group 文件夹数目
        /// </summary>
        public static int GetNextMaxGroup()
        {
            string[] directories = Directory.GetDirectories(WinXFolder, "Group*", SearchOption.TopDirectoryOnly);
            Array.Sort(directories, ReverseSort);

            foreach (string path in directories)
            {
                string ss = Path.GetFileName(path)?.Substring(5).Trim();
                if (ss == null) { continue; }

                int result;
                if (int.TryParse(ss, out result)) { return result + 1; }
            }
            return 0;
        }

        /// <summary>
        /// 转换环境变量为绝对路径
        /// </summary>
        private static string ExpandEnvVar(string withVars)
        {
            if (withVars == null) { return null; }
            MatchCollection matches = Regex.Matches(withVars, "%.*%");

            foreach (Match match in matches)
            {
                if (Environment.GetEnvironmentVariable(match.Value.Trim('%')) == null)
                {
                    return null;
                }
            }
            return Environment.ExpandEnvironmentVariables(withVars);
        }

        /// <summary>
        /// 快捷方式提权操作
        /// </summary>
        private static void ElevateLnk(string lnkFile)
        {
            /*
             * 作者：Henrik Rading
             * Blog：https://blog.ctglobalservices.com/powershell/hra/create-shortcut-with-elevated-rights/
             */

            using (FileStream fileStream = new FileStream(lnkFile, FileMode.Open, FileAccess.ReadWrite))
            {
                fileStream.Seek(21, SeekOrigin.Begin);
                fileStream.WriteByte(0x22);
            }
        }

        private static string _groupFolderPath;

        private class ReverseSorter : System.Collections.IComparer
        {
            int System.Collections.IComparer.Compare(object x, object y)
            {
                return new System.Collections.CaseInsensitiveComparer().Compare(y, x);
            }
        }

        private static readonly System.Collections.IComparer ReverseSort = new ReverseSorter();

        private static readonly WshShell wshShell = new WshShell();

        private static readonly KeyValuePair<string, string>[] SystemFolderMapping = new[]
        {
            new KeyValuePair<string, string>(ExpandEnvVar("%PROGRAMFILES%").ToLower(), "{905E63B6-C1BF-494E-B29C-65B732D3D21A}"),
            new KeyValuePair<string, string>(ExpandEnvVar("%WINDIR%\\System32").ToLower(), "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"),
            new KeyValuePair<string, string>(ExpandEnvVar("%WINDIR%").ToLower(), "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"),
            new KeyValuePair<string, string>(ExpandEnvVar("%SYSTEMROOT%").ToLower(), "{F38BF404-1D43-42F2-9305-67DE0B28FC23}")
        };

        private static readonly string WinXFolder = ExpandEnvVar(@"%LOCALAPPDATA%\Microsoft\Windows\WinX");

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern int HashData(
            [In, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 1)] byte[] pbData,
            int cbData,
            [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 3)] byte[] pbHash,
            int cbHash);
    }
}
