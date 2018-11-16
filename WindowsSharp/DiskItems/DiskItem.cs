﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
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

                        return SubItemsList;
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

        public String ItemRealName
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
                    string outPath = Path.GetFileName(ItemPath);
                    if (String.IsNullOrWhiteSpace(outPath))
                        return ItemPath;
                    else
                        return outPath;
                }
            }
        }

        string _itemDisplayName = null;

        public String ItemDisplayName
        {
            get
            {
                if (_itemDisplayName != null)
                    return _itemDisplayName;
                else
                    return ItemRealName;
            }
            set
            {
                _itemDisplayName = value;
            }
        }

        public System.Drawing.Icon ItemSmallIcon
        {
            get
            {
                if ((File.Exists(ItemPath)) || (Directory.Exists(ItemPath)))
                {
                    UInt32 flags = (0x00000001 | 0x100);
                    NativeMethods.ShFileInfo shInfo = new NativeMethods.ShFileInfo();
                    NativeMethods.SHGetFileInfo(ItemPath, 0, ref shInfo, (UInt32)Marshal.SizeOf(shInfo), flags);
                    //return System.Drawing.Icon.FromHandle(shInfo.hIcon);
                    System.Drawing.Icon result = (System.Drawing.Icon)(System.Drawing.Icon.FromHandle(shInfo.hIcon).Clone());
                    NativeMethods.DestroyIcon(shInfo.hIcon);
                    return result;
                }
                else return null;
            }
        }

        public System.Drawing.Icon ItemLargeIcon
        {
            get
            {
                if ((File.Exists(ItemPath)) || (Directory.Exists(ItemPath)))
                {
                    UInt32 flags = (0x000000000 | 0x100);
                    NativeMethods.ShFileInfo shInfo = new NativeMethods.ShFileInfo();
                    NativeMethods.SHGetFileInfo(ItemPath, 0, ref shInfo, (UInt32)Marshal.SizeOf(shInfo), flags);
                    System.Drawing.Icon result = (System.Drawing.Icon)(System.Drawing.Icon.FromHandle(shInfo.hIcon).Clone());
                    NativeMethods.DestroyIcon(shInfo.hIcon);
                    return result;
                }
                else return null;
            }
        }

        public System.Drawing.Icon ItemExtraLargeIcon
        {
            get
            {
                if ((File.Exists(ItemPath)) || (Directory.Exists(ItemPath)))
                {
                    UInt32 flags = (0x000000000 | 0x100);
                    NativeMethods.ShFileInfo shInfo = new NativeMethods.ShFileInfo();
                    NativeMethods.SHGetFileInfo(ItemPath, 0, ref shInfo, (UInt32)Marshal.SizeOf(shInfo), flags);
                    System.Drawing.Icon result = (System.Drawing.Icon)(System.Drawing.Icon.FromHandle(shInfo.hIcon).Clone());
                    NativeMethods.DestroyIcon(shInfo.hIcon);

                    NativeMethods.IImageList list;
                    var hres = NativeMethods.SHGetImageList(0x2, ref NativeMethods.iidImageList, out list);
                    IntPtr resultHandle = IntPtr.Zero;
                    list.GetIcon(shInfo.iIcon, 1, ref resultHandle);
                    System.Drawing.Icon finalResult = (System.Drawing.Icon)(System.Drawing.Icon.FromHandle(resultHandle).Clone());
                    NativeMethods.DestroyIcon(resultHandle);
                    return finalResult;
                }
                else return null;
            }
        }

        public System.Drawing.Icon ItemJumboIcon
        {
            get
            {
                if ((File.Exists(ItemPath)) || (Directory.Exists(ItemPath)))
                {
                    UInt32 flags = (0x000000000 | 0x100);
                    NativeMethods.ShFileInfo shInfo = new NativeMethods.ShFileInfo();
                    NativeMethods.SHGetFileInfo(ItemPath, 0, ref shInfo, (UInt32)Marshal.SizeOf(shInfo), flags);
                    System.Drawing.Icon result = (System.Drawing.Icon)(System.Drawing.Icon.FromHandle(shInfo.hIcon).Clone());
                    NativeMethods.DestroyIcon(shInfo.hIcon);

                    NativeMethods.IImageList list;
                    var hres = NativeMethods.SHGetImageList(0x4, ref NativeMethods.iidImageList, out list);
                    IntPtr resultHandle = IntPtr.Zero;
                    list.GetIcon(shInfo.iIcon, 1, ref resultHandle);
                    System.Drawing.Icon finalResult = (System.Drawing.Icon)(System.Drawing.Icon.FromHandle(resultHandle).Clone());
                    NativeMethods.DestroyIcon(resultHandle);
                    return finalResult;
                }
                else return null;
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
                        //Debug.WriteLine("dir " + s);
                        items.Add(new DiskItem(s));
                    }
                    foreach (var s in Directory.EnumerateFiles(ItemPath))
                    {
                        //Debug.WriteLine("file " + s);
                        items.Add(new DiskItem(s));
                    }
                }
                return items;
            }
        }

        public double ItemSize
        {
            get
            {
                if ((ItemCategory == DiskItemCategory.File) || (ItemCategory == DiskItemCategory.Shortcut))
                {
                    return new FileInfo(_itemPath).Length;
                }
                else return (double)0.0;
            }
        }

        public string FriendlyItemSize
        {
            get
            {
                if ((ItemCategory == DiskItemCategory.File) || (ItemCategory == DiskItemCategory.Shortcut))
                {
                    double size = ItemSize;
                    int unitCounter = 0;
                    while (size > 1024)
                    {
                        size = size / 1024;
                        unitCounter++;
                    }

                    if (unitCounter == 0)
                        return size.ToString() + " B";
                    else if (unitCounter == 1)
                        return size.ToString() + " KB";
                    else if (unitCounter == 2)
                        return size.ToString() + " MB";
                    else if (unitCounter == 3)
                        return size.ToString() + " GB";
                    else if (unitCounter == 4)
                        return size.ToString() + " TB";
                    else
                        return size.ToString() + " PB";
                }
                else return String.Empty;
            }
        }

        NativeMethods.ShFileInfo _fileInfo = new NativeMethods.ShFileInfo();

        public DiskItem(String path)
        {
            ItemPath = Environment.ExpandEnvironmentVariables(path);
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
            else if (Directory.Exists(ItemPath))
                ItemCategory = DiskItemCategory.Directory;
            else// if (!(Directory.Exists(ItemPath)))
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

        public enum OpenVerbs
        {
            Normal,
            Admin
        }

        public void Open()
        {
            Open(OpenVerbs.Normal, string.Empty);
        }

        public void Open(string args)
        {
            Open(OpenVerbs.Normal, args);
        }

        public void Open(OpenVerbs verb)
        {
            Open(verb, string.Empty);
        }

        public void Open(OpenVerbs verb, string args)
        {
            if (ItemCategory == DiskItemCategory.App)
            {
                LaunchApp(ItemRealName);
            }
            else
            {
                if (verb == OpenVerbs.Admin)
                    Process.Start(new ProcessStartInfo(_itemPath, args)
                    {
                        Verb = "runas"
                    });
                else
                    Process.Start(_itemPath, args);
            }
        }

        uint LaunchApp(string packageFullName, string arguments = null)
        {
            try
            {
                IntPtr pir = IntPtr.Zero;
                NativeMethods.OpenPackageInfoByFullName(packageFullName, 0, out pir);

                int length = 0, count;
                NativeMethods.GetPackageApplicationIds(pir, ref length, null, out count);

                var buffer = new byte[length];
                NativeMethods.GetPackageApplicationIds(pir, ref length, buffer, out count);

                var appUserModelId = Encoding.Unicode.GetString(buffer, IntPtr.Size * count, length - IntPtr.Size * count);

                var activation = (NativeMethods.IApplicationActivationManager)new NativeMethods.ApplicationActivationManager();
                uint pid;
                int hr = activation.ActivateApplication(appUserModelId, arguments ?? string.Empty, NativeMethods.ActivateOptions.None, out pid);
                if (hr < 0)
                    Marshal.ThrowExceptionForHR(hr);
                return pid;
            }
            catch
            {
                Process.Start(packageFullName);
                return 0;
            }
        }

        public bool ShowProperties()
        {
            /*NativeMethods.SHELLEXECUTEINFO info = new NativeMethods.SHELLEXECUTEINFO()
            {
                lpVerb = "properties",
                lpFile = "\"" + _itemPath + "\"",
                nShow = 5,
                fMask = 12,
            };
            info.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(info);
            return NativeMethods.ShellExecuteEx(ref info);*/
            bool returnValue = false;
            try
            {
                var info = new ProcessStartInfo(_itemPath)
                {
                    Verb = "properties",
                    UseShellExecute = true
                };
                returnValue = (Process.Start(info) != null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("1 " + ex);
                try
                {
                    var info = new NativeMethods.SHELLEXECUTEINFO()
                    {
                        lpFile = _itemPath,
                        lpVerb = "properties"
                    };
                    return NativeMethods.ShellExecuteEx(ref info);
                }
                catch (Exception exc)
                {
                    Debug.WriteLine("2 " + exc);
                    return false;
                }
            }

            if (!returnValue)
            {
                try
                {
                    var info = new NativeMethods.SHELLEXECUTEINFO()
                    {
                        lpFile = _itemPath,
                        lpVerb = "properties",
                        nShow = 5,
                        fMask = 0x0000000C
                    };
                    //info.cbSize = sizeof(info);
                    return NativeMethods.ShellExecuteEx(ref info);
                }
                catch (Exception exce)
                {
                    Debug.WriteLine("3 " + exce);
                    return false;
                }
            }

            return returnValue;
        }

        private class NativeMethods
        {
            [DllImport("shell32.dll", EntryPoint = "#727")]
            public extern static int SHGetImageList(int iImageList, ref Guid riid, out IImageList ppv);

            public static Guid iidImageList = new Guid("46EB5926-582E-4017-9FDF-E8998DAA0950");

            [ComImportAttribute()]
            [GuidAttribute("46EB5926-582E-4017-9FDF-E8998DAA0950")]
            [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
            //helpstring("Image List"),
            public interface IImageList
            {
                [PreserveSig]
                int Add(
                    IntPtr hbmImage,
                    IntPtr hbmMask,
                    ref int pi);

                [PreserveSig]
                int ReplaceIcon(
                    int i,
                    IntPtr hicon,
                    ref int pi);

                [PreserveSig]
                int SetOverlayImage(
                    int iImage,
                    int iOverlay);

                [PreserveSig]
                int Replace(
                    int i,
                    IntPtr hbmImage,
                    IntPtr hbmMask);

                [PreserveSig]
                int AddMasked(
                    IntPtr hbmImage,
                    int crMask,
                    ref int pi);

                [PreserveSig]
                int Draw(
                    ref IMAGELISTDRAWPARAMS pimldp);

                [PreserveSig]
                int Remove(
                int i);

                [PreserveSig]
                int GetIcon(
                    int i,
                    int flags,
                    ref IntPtr picon);

                [PreserveSig]
                int GetImageInfo(
                    int i,
                    ref IMAGEINFO pImageInfo);

                [PreserveSig]
                int Copy(
                    int iDst,
                    IImageList punkSrc,
                    int iSrc,
                    int uFlags);

                [PreserveSig]
                int Merge(
                    int i1,
                    IImageList punk2,
                    int i2,
                    int dx,
                    int dy,
                    ref Guid riid,
                    ref IntPtr ppv);

                [PreserveSig]
                int Clone(
                    ref Guid riid,
                    ref IntPtr ppv);

                [PreserveSig]
                int GetImageRect(
                    int i,
                    ref System.Drawing.Rectangle prc);

                [PreserveSig]
                int GetIconSize(
                    ref int cx,
                    ref int cy);

                [PreserveSig]
                int SetIconSize(
                    int cx,
                    int cy);

                [PreserveSig]
                int GetImageCount(
                ref int pi);

                [PreserveSig]
                int SetImageCount(
                    int uNewCount);

                [PreserveSig]
                int SetBkColor(
                    int clrBk,
                    ref int pclr);

                [PreserveSig]
                int GetBkColor(
                    ref int pclr);

                [PreserveSig]
                int BeginDrag(
                    int iTrack,
                    int dxHotspot,
                    int dyHotspot);

                [PreserveSig]
                int EndDrag();

                [PreserveSig]
                int DragEnter(
                    IntPtr hwndLock,
                    int x,
                    int y);

                [PreserveSig]
                int DragLeave(
                    IntPtr hwndLock);

                [PreserveSig]
                int DragMove(
                    int x,
                    int y);

                [PreserveSig]
                int SetDragCursorImage(
                    ref IImageList punk,
                    int iDrag,
                    int dxHotspot,
                    int dyHotspot);

                [PreserveSig]
                int DragShowNolock(
                    int fShow);

                [PreserveSig]
                int GetDragImage(
                    ref System.Drawing.Point ppt,
                    ref System.Drawing.Point pptHotspot,
                    ref Guid riid,
                    ref IntPtr ppv);

                [PreserveSig]
                int GetItemFlags(
                    int i,
                    ref int dwFlags);

                [PreserveSig]
                int GetOverlayImage(
                    int iOverlay,
                    ref int piIndex);
            };

            [StructLayout(LayoutKind.Sequential)]
            public struct IMAGEINFO
            {
                public IntPtr hbmImage;
                public IntPtr hbmMask;
                public int Unused1;
                public int Unused2;
                public System.Drawing.Rectangle rcImage;
            }

            public struct IMAGELISTDRAWPARAMS
            {
                public int cbSize;
                public IntPtr himl;
                public int i;
                public IntPtr hdcDst;
                public int x;
                public int y;
                public int cx;
                public int cy;
                public int xBitmap;        // x offest from the upperleft of bitmap
                public int yBitmap;        // y offset from the upperleft of bitmap
                public int rgbBk;
                public int rgbFg;
                public int fStyle;
                public int dwRop;
                public int fState;
                public int Frame;
                public int crEffect;
            }


            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool DestroyIcon(IntPtr hIcon);

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

            [DllImport("shell32.dll")]
            public static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]

            public struct SHELLEXECUTEINFO
            {
                public int cbSize;
                public uint fMask;
                public IntPtr hwnd;
                public String lpVerb;
                public String lpFile;
                public String lpParameters;
                public String lpDirectory;
                public int nShow;
                public int hInstApp;
                public int lpIDList;
                public String lpClass;
                public int hkeyClass;
                public uint dwHotKey;
                public int hIcon;
                public int hProcess;
            }

            public enum ActivateOptions
            {
                None = 0x00000000,  // No flags set
                DesignMode = 0x00000001,  // The application is being activated for design mode
                NoErrorUI = 0x00000002,  // Do not show an error dialog if the app fails to activate                                
                NoSplashScreen = 0x00000004,  // Do not show the splash screen when activating the app
            }

            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("2e941141-7f97-4756-ba1d-9decde894a3d")]
            public interface IApplicationActivationManager
            {
                int ActivateApplication([MarshalAs(UnmanagedType.LPWStr)] string appUserModelId, [MarshalAs(UnmanagedType.LPWStr)] string arguments,
                    ActivateOptions options, out uint processId);
                int ActivateForFile([MarshalAs(UnmanagedType.LPWStr)] string appUserModelId, IntPtr pShelItemArray,
                    [MarshalAs(UnmanagedType.LPWStr)] string verb, out uint processId);
                int ActivateForProtocol([MarshalAs(UnmanagedType.LPWStr)] string appUserModelId, IntPtr pShelItemArray,
                    [MarshalAs(UnmanagedType.LPWStr)] string verb, out uint processId);
            }

            [DllImport("kernel32")]
            public static extern int OpenPackageInfoByFullName([MarshalAs(UnmanagedType.LPWStr)] string fullName, uint reserved, out IntPtr packageInfo);

            [DllImport("kernel32")]
            public static extern int GetPackageApplicationIds(IntPtr pir, ref int bufferLength, byte[] buffer, out int count);

            [DllImport("kernel32")]
            public static extern int ClosePackageInfo(IntPtr pir);

            [ComImport, Guid("45BA127D-10A8-46EA-8AB7-56EA9078943C")]
            public class ApplicationActivationManager { }
        }
    }
}