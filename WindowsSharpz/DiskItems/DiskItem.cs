using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsSharp.DiskItems
{
    public class DiskItem
    {
        public static List<DiskItem> AllApps
        {
            get
            {
                List<DiskItem> items = new List<DiskItem>();

                List<DiskItem> AllAppsAppDataItems = new List<DiskItem>();
                foreach (var s in Directory.EnumerateFiles(Environment.ExpandEnvironmentVariables(@"%appdata%\Microsoft\Windows\Start Menu\Programs")))
                {
                    AllAppsAppDataItems.Add(new DiskItem(s));
                }
                foreach (var s in Directory.EnumerateDirectories(Environment.ExpandEnvironmentVariables(@"%appdata%\Microsoft\Windows\Start Menu\Programs")))
                {
                    AllAppsAppDataItems.Add(new DiskItem(s));
                }

                List<DiskItem> AllAppsProgramDataItems = new List<DiskItem>();
                foreach (var s in Directory.EnumerateFiles(Environment.ExpandEnvironmentVariables(@"%programdata%\Microsoft\Windows\Start Menu\Programs")))
                {
                    AllAppsProgramDataItems.Add(new DiskItem(s));
                }
                foreach (var s in Directory.EnumerateDirectories(Environment.ExpandEnvironmentVariables(@"%programdata%\Microsoft\Windows\Start Menu\Programs")))
                {
                    AllAppsProgramDataItems.Add(new DiskItem(s));
                }

                List<DiskItem> AllAppsItems = new List<DiskItem>();
                List<DiskItem> AllAppsReorgItems = new List<DiskItem>();
                foreach (DiskItem t in AllAppsAppDataItems)
                {
                    var FolderIsDuplicate = false;

                    foreach (DiskItem v in AllAppsProgramDataItems)
                    {
                        List<DiskItem> SubItemsList = new List<DiskItem>();

                        if (Directory.Exists(t.ItemPath))
                        {
                            if (((t.ItemCategory == DiskItemCategory.Directory) & (v.ItemCategory == DiskItemCategory.Directory)) && ((t.ItemPath.Substring(t.ItemPath.LastIndexOf(@"\"))) == (v.ItemPath.Substring(v.ItemPath.LastIndexOf(@"\")))))
                            {
                                FolderIsDuplicate = true;
                                foreach (var i in Directory.EnumerateDirectories(t.ItemPath))
                                {
                                    SubItemsList.Add(new DiskItem(i));
                                }

                                foreach (var j in Directory.EnumerateFiles(v.ItemPath))
                                {
                                    SubItemsList.Add(new DiskItem(j));
                                }
                            }
                        }

                        if (!AllAppsItems.Contains(v))
                        {
                            AllAppsItems.Add(v);
                        }
                    }

                    if ((!AllAppsItems.Contains(t)) && (!FolderIsDuplicate))
                    {
                        AllAppsItems.Add(t);
                    }
                }

                foreach (DiskItem x in AllAppsItems)
                {
                    if (File.Exists(x.ItemPath))
                    {
                        AllAppsReorgItems.Add(x);
                    }
                }

                foreach (DiskItem x in AllAppsItems)
                {
                    if (Directory.Exists(x.ItemPath))
                    {
                        AllAppsReorgItems.Add(x);
                    }
                }

                return AllAppsReorgItems;
            }
        }

        public String ItemName
        {
            get
            {
                if (ItemCategory == DiskItemCategory.App)
                    return ItemAppInfo.DisplayName;
                else if (ItemCategory == DiskItemCategory.Shortcut)
                    return Path.GetFileName(ItemPath); //Temporary maybe?
                else
                {
                    /*if (Path.GetExtension(ItemPath).ToLower().EndsWith("exe"))
                        try
                        {
                            return new FileInfo(ItemPath).;
                            return System.Reflection.Assembly.LoadFile(ItemPath).FullName;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex);
                            return Path.GetFileName(ItemPath);
                        }
                    else*/
                    return Path.GetFileName(ItemPath);
                }
            }
        }

        String _itemPath;

        public String ItemPath
        {
            get
            {
                if (ItemCategory == DiskItemCategory.Shortcut)
                {
                    var raw = _itemPath;

                    var targetPath = GetMsiShortcut(raw);

                    if (targetPath == null)
                    {
                        targetPath = ResolveShortcut(raw);
                    }

                    if (targetPath == null)
                    {
                        targetPath = GetInternetShortcut(raw);
                    }

                    if (targetPath == null | targetPath == "" | targetPath.Replace(" ", "") == "")
                    {
                        return raw;
                    }
                    else
                    {
                        return targetPath;
                    }
                }
                else return _itemPath;
            }
            set => _itemPath = value;
        }

        String GetInternetShortcut(String _rawPath)
        {
            try
            {
                var url = "";

                using (TextReader reader = new StreamReader(_rawPath))
                {
                    var line = "";
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.StartsWith("URL="))
                        {
                            String[] splitLine = line.Split('=');
                            if (splitLine.Length > 0)
                            {
                                url = splitLine[1];
                                break;
                            }
                        }
                    }
                }
                return url;
            }
            catch
            {
                return null;
            }
        }

        String ResolveShortcut(String _rawPath)
        {
            // IWshRuntimeLibrary is in the COM library "Windows Script Host Object Model"
            /*IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();

            try
            {
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(_rawPath);
                return shortcut.TargetPath;
            }
            catch
            {*/
                // A COMException is thrown if the file is not a valid shortcut (.lnk) file 
                return null;
            /*}*/
        }

        String GetMsiShortcut(String _rawPath)
        {
            var product = new StringBuilder(NativeMethods.MaxGuidLength + 1);
            var feature = new StringBuilder(NativeMethods.MaxFeatureLength + 1);
            var component = new StringBuilder(NativeMethods.MaxGuidLength + 1);

            NativeMethods.MsiGetShortcutTarget(_rawPath, product, feature, component);

            var pathLength = NativeMethods.MaxPathLength;
            var path = new StringBuilder(pathLength);

            NativeMethods.InstallState installState = NativeMethods.MsiGetComponentPath(product.ToString(), component.ToString(), path, ref pathLength);
            if (installState == NativeMethods.InstallState.Local)
            {
                return path.ToString();
            }
            else
            {
                return null;
            }
        }

        public AppInfo ItemAppInfo { get; set; }

        public enum DiskItemCategory
        {
            File,
            Shortcut,
            Directory,
            App
        }

        public DiskItemCategory ItemCategory { get; set; }

        public String FriendlyItemType { get; set; }

        public List<DiskItem> SubItems
        {
            get
            {
                List<DiskItem> items = new List<DiskItem>();
                if (ItemCategory == DiskItemCategory.Directory)
                {
                    foreach (var s in Directory.EnumerateDirectories(ItemPath))
                    {
                        items.Add(new DiskItem(s));
                    }
                    foreach (var s in Directory.EnumerateFiles(ItemPath))
                    {
                        items.Add(new DiskItem(s));
                    }
                }
                return items;
            }
        }

        NativeMethods.ShFileInfo _fileInfo = new NativeMethods.ShFileInfo();

        public DiskItem(String path)
        {
            ItemPath = path;
            if (File.Exists(ItemPath))
            {
                if (Path.GetExtension(path).EndsWith("lnk"))
                {
                    ItemCategory = DiskItemCategory.Shortcut;
                }
                else
                {
                    ItemCategory = DiskItemCategory.File;
                }
            }
            else if (!(Directory.Exists(ItemPath)))
            {
                if (Directory.Exists(Environment.ExpandEnvironmentVariables(@"%programfiles%\WindowsApps\" + path)))
                {
                    ItemCategory = DiskItemCategory.App;
                    ItemAppInfo = new AppInfo(path);
                }
            }

            if (ItemCategory == DiskItemCategory.File && File.Exists(path))
            {

                if (string.IsNullOrEmpty(_fileInfo.szTypeName))
                {
                    FriendlyItemType = Path.GetExtension(ItemPath).Replace(".", "").ToUpper() + " File";
                }
                else
                {
                    FriendlyItemType = _fileInfo.szTypeName;
                }
            }
            else if (ItemCategory == DiskItemCategory.Shortcut)
            {
                FriendlyItemType = "Shortcut";
            }
            else if (ItemCategory == DiskItemCategory.Directory)// Directory.Exists(ItemPath))
            {
                FriendlyItemType = "File Folder";
            }
            else if (ItemCategory == DiskItemCategory.App)
            {
                if (Environment.OSVersion.Version.Major >= 10)
                {
                    FriendlyItemType = "Universal App";
                }
                else
                {
                    FriendlyItemType = "Modern App";
                }
            }
        }

        private class NativeMethods
        {
            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
            public static extern IntPtr SHGetFileInfo(String pszPath, UInt32 dwFileAttributes, ref ShFileInfo psfi, UInt32 cbFileInfo, UInt32 uFlags);

            [StructLayout(LayoutKind.Sequential)]
            public struct ShFileInfo
            {
                public IntPtr hIcon;
                public Int32 iIcon;
                public UInt32 dwAttributes;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public String szDisplayName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
                public String szTypeName;
            }

            [DllImport("msi.dll", CharSet = CharSet.Auto)]
            public static extern UInt32 MsiGetShortcutTarget(String targetFile, StringBuilder productCode, StringBuilder featureID, StringBuilder componentCode);

            public const Int32 MaxFeatureLength = 38;
            public const Int32 MaxGuidLength = 38;
            public const Int32 MaxPathLength = 1024;

            public enum InstallState
            {
                NotUsed = -7,
                BadConfig = -6,
                Incomplete = -5,
                SourceAbsent = -4,
                MoreData = -3,
                InvalidArg = -2,
                Unknown = -1,
                Broken = 0,
                Advertised = 1,
                Removed = 1,
                Absent = 2,
                Local = 3,
                Source = 4,
                Default = 5
            }

            [DllImport("msi.dll", CharSet = CharSet.Auto)]
            public static extern InstallState MsiGetComponentPath(String productCode, String componentCode, StringBuilder componentPath, ref Int32 componentPathBufferSize);
        }
    }
}