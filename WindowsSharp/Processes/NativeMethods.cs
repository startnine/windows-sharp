using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Text;

namespace WindowsSharp
{
    public static class NativeMethods
    {
        //Begin tray stuff ( https://gist.github.com/paulcbetts/e90c7d89624d1b1adc72 )

        [StructLayout(LayoutKind.Sequential)]
        public struct GUID
        {
            public int a;
            public short b;
            public short c;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] d;
        }

        // NOTIFYITEM describes an entry in Explorer's registry of status icons.
        // Explorer keeps entries around for a process even after it exits.
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct NOTIFYITEM
        {
            /*[MarshalAs(UnmanagedType.LPWStr)]
            public string exe_name;    // The file name of the creating executable.

            [MarshalAs(UnmanagedType.LPWStr)]
            public string tip;         // The last hover-text value associated with this status
                                // item.

            public IntPtr icon;       // The icon associated with this status item.
            public IntPtr hwnd;       // The HWND associated with the status item.
            public int preference;  // Determines the behavior of the icon with respect to
                                                      // the taskbar
            public uint id;           // The ID specified by the application.  (hWnd, uID) is
                                      // unique.
            public Guid guid;         // The GUID specified by the application, alternative to
                                      // uID.

            public uint uCallbackMessage;*/

            public string pszExeName { get; set; }
            public string pszTip { get; set; }
            public IntPtr hIcon { get; set; }
            public IntPtr hWnd { get; set; }
            public uint dwPreference { get; set; }
            public uint uID { get; set; }
            public GUID GuidItem { get; set; }
            public uint uCallbackMessage { get; set; }
            /*public uint reserved { get; set; }
            public uint reserved2 { get; set; }
            public uint reserved3 { get; set; }
            public uint reserved4 { get; set; }
            public uint reserved5 { get; set; }*/
            void modify(NOTIFYICONDATA pnid)
            {
                bool change = false;
            }
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public class NOTIFYICONDATA
        {
            public int cbSize = Marshal.SizeOf(typeof(NOTIFYICONDATA));
            public IntPtr hWnd;
            public int uID;
            public int uFlags;
            public int uCallbackMessage;
            public IntPtr hIcon;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x80)]
            public string szTip;
            public int dwState;
            public int dwStateMask;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x100)]
            public string szInfo;
            public int uTimeoutOrVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x40)]
            public string szInfoTitle;
            public int dwInfoFlags;
        }

        [ComImport]
        [Guid("D782CCBA-AFB0-43F1-94DB-FDA3779EACCB")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface INotificationCb
        {
            void Notify([In]uint nEvent, [In] ref NOTIFYITEM notifyItem);
            //IntPtr QueryInterface(int refiId, object idk);
        }

        [ComImport]
        [Guid("FB852B2C-6BAD-4605-9551-F15F87830935")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ITrayNotifyWin7
        {
            void RegisterCallback([MarshalAs(UnmanagedType.Interface)]INotificationCb callback);
            void SetPreference([In] ref NOTIFYITEM notifyItem);
            void EnableAutoTray([In] bool enabled);
        }

        [ComImport]
        [Guid("D133CE13-3537-48BA-93A7-AFCD5D2053B4")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ITrayNotify
        {
            void RegisterCallback([MarshalAs(UnmanagedType.Interface)]INotificationCb callback, [Out] out ulong handle);
            void UnregisterCallback([In] ref ulong handle);
            void SetPreference([In] ref NOTIFYITEM notifyItem);
            void EnableAutoTray([In] bool enabled);
            void DoAction([In] bool enabled);
        }

        [ComImport, Guid("25DEAD04-1EAC-4911-9E3A-AD0A4AB560FD")]
        public class TrayNotify { }

        public class NotificationCb : INotificationCb
        {
            public readonly List<NOTIFYITEM> items = new List<NOTIFYITEM>();

            public void Notify([In] uint nEvent, [In] ref NOTIFYITEM notifyItem)
            {
                items.Add(notifyItem);
            }
        }
        //End tray stuff


        [DllImport("dwmapi.dll")]
        static extern Int32 DwmIsCompositionEnabled(out Boolean enabled);

        public static bool DwmIsCompositionEnabled()
        {
            DwmIsCompositionEnabled(out bool returnValue);
            return returnValue;
        }

        [DllImport("gdi32.dll")]
        public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hObjectSource, int nXSrc, int nYSrc, int dwRop);
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth, int nHeight);
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
        [DllImport("gdi32.dll")]
        public static extern bool DeleteDC(IntPtr hDC);
        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);
        [DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);


        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);
        [DllImport("user32.dll")]
        public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);


        [DllImport("dwmapi.dll", EntryPoint = "#113", SetLastError = true)]
        public static extern uint DwmpActivateLivePreview(uint peekOn, IntPtr hWnd, IntPtr hTopmostWindow, uint peekType1or3);

        [DllImport("dwmapi.dll", EntryPoint = "#113", SetLastError = true)]
        public static extern uint DwmpActivateLivePreview(uint peekOn, IntPtr hWnd, IntPtr hTopmostWindow, uint peekType1or3, UIntPtr passZeroHere);

        //DwmpActivateLivePreview(1, Handle, topmostWindowHandle, 1);//activate
        //DwmpActivateLivePreview(0, Handle, topmostWindowHandle, 1);//deactivate

        [DllImport("psapi.dll")]
        public static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, [Out] StringBuilder lpBaseName, [In] [MarshalAs(UnmanagedType.U4)] int nSize);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool QueryFullProcessImageName(IntPtr hProcess, int dwFlags, StringBuilder lpExeName, out int lpdwSize);
        /*[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool QueryFullProcessImageName(IntPtr hProcess, uint dwFlags, [Out, MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpExeName, ref uint lpdwSize);*/

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(int flags, bool inherit, int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
        [SuppressUnmanagedCodeSecurity]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hObject);

        [DllImport("User32.dll")]
        public static extern bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        public delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        public static Int32 WmSetText = 0x000C;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll")]
        public static extern uint TrackPopupMenuEx(IntPtr hmenu, uint fuFlags, int x, int y, IntPtr hwnd, IntPtr lptpm);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        internal struct WINDOWPLACEMENT
        {
            /// <summary>
            /// The length of the structure, in bytes. Before calling the GetWindowPlacement or SetWindowPlacement functions, set this member to sizeof(WINDOWPLACEMENT).
            /// <para>
            /// GetWindowPlacement and SetWindowPlacement fail if this member is not set correctly.
            /// </para>
            /// </summary>
            public int Length;

            /// <summary>
            /// Specifies flags that control the position of the minimized window and the method by which the window is restored.
            /// </summary>
            public int Flags;

            /// <summary>
            /// The current show state of the window.
            /// </summary>
            public Int32 ShowCmd;

            /// <summary>
            /// The coordinates of the window's upper-left corner when the window is minimized.
            /// </summary>
            public Point MinPosition;

            /// <summary>
            /// The coordinates of the window's upper-left corner when the window is maximized.
            /// </summary>
            public Point MaxPosition;

            /// <summary>
            /// The window's coordinates when the window is in the restored position.
            /// </summary>
            public Rectangle NormalPosition;

            /// <summary>
            /// Gets the default (empty) value.
            /// </summary>
            public static WINDOWPLACEMENT Default
            {
                get
                {
                    WINDOWPLACEMENT result = new WINDOWPLACEMENT();
                    result.Length = Marshal.SizeOf(result);
                    return result;
                }
            }
        }

        [DllImport("user32.dll")]
        public static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpfn, IntPtr lParam);

        public delegate Boolean EnumDelegate(IntPtr hWnd, Int32 lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern Int32 GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern void GetWindowText(IntPtr hWnd, StringBuilder lpString, Int32 nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern Boolean IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public static uint WmClose = 0x0010;
        public static uint WmMouseMove = 0x0200;
        public static uint WmLButtonDown = 0x0201;
        public static uint WmMButtonDown = 0x0207;
        public static uint WmRButtonDown = 0x0204;
        public static uint WmRButtonUp = 0x0205;
        public static uint WmXButtonDown = 0x020B;

        [DllImport("user32.dll")]
        public static extern Boolean SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32")]
        public static extern Int32 ShowWindow(IntPtr hwnd, Int32 nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hwnd, Int32 nCmdShow);

        public static int SwForceMinimize = 11;
        public static int SwHide = 0;
        public static int SwMaximize = 3;
        public static int SwMinimize = 6;
        public static int SwRestore = 9;
        public static int SwShow = 5;
        public static int SwShowDefault = 10;
        public static int SwShowMaximized = 3;
        public static int SwShowMinimized = 2;
        public static int SwShowMinimizedNoActive = 7;
        public static int SwShowNa = 8;
        public static int SwShowNoActivate = 4;
        public static int SwShowNormal = 1;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern Boolean IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "RegisterWindowMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int RegisterWindowMessage(string lpString);

        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int DeregisterShellHookWindow(IntPtr hWnd);

        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int RegisterShellHookWindow(IntPtr hWnd);

        public static int HShellWindowCreated = 1;
        public static int HShellWindowDestroyed = 2;
        public static int HShellActivateShellWindow = 3;
        public static int HShellWindowActivated = 4;
        public static int HShellGetMinRect = 5;
        public static int HShellRedraw = 6;
        public static int HShellTaskMan = 7;
        public static int HShellLanguage = 8;
        public static int HShellAccessibilityState = 11;
        public static int HShellAppCommand = 12;

        public static IntPtr GetClassLongPtr(IntPtr hWnd, int nIndex) => (!Environment.Is64BitOperatingSystem)
        ? GetClassLongPtr64(hWnd, nIndex)
        : GetClassLongPtr32(hWnd, nIndex);
        /*{
            if (IntPtr.Size > 4)
                return GetClassLongPtr64(hWnd, nIndex);
            else
                return new IntPtr(GetClassLongPtr32(hWnd, nIndex));
        }*/

        [DllImport("user32.dll", EntryPoint = "GetClassLong")]
        public static extern IntPtr GetClassLongPtr32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetClassLongPtr")]
        public static extern IntPtr GetClassLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, UIntPtr lParam);

        public const int GclHIconSm = -34;
        public const int GclHIcon = -14;

        public const int IconSmall = 0;
        public const int IconBig = 1;
        public const int IconSmall2 = 2;

        public const int WmGetIcon = 0x7F;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern Boolean DestroyIcon(IntPtr hIcon);

        public static IntPtr GetWindowLong(IntPtr hWnd, Int32 nIndex) => (!Environment.Is64BitOperatingSystem)
        ? GetWindowLong64(hWnd, nIndex)
        : GetWindowLong32(hWnd, nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr GetWindowLong32(IntPtr window, int offset);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        static extern IntPtr GetWindowLong64(IntPtr hWnd, Int32 nIndex);

        public const Int32 GwlStyle = -16;

        public const Int32 WsVisible = 0x10000000;
        public const Int32 WsBorder = 0x00800000;
        public const Int32 WsChild = 0x40000000;

        public const Int32 GwlExStyle = -20;

        public const Int32 WsExToolWindow = 0x00000080;

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    }

    ////https://stackoverflow.com/questions/32122679/getting-icon-of-modern-windows-app-from-a-desktop-application/36495656
    public static class AppxMethods
    {
        /*static void Main(string[] args)
        {
            foreach (var p in Process.GetProcesses())
            {
                var package = AppxPackage.FromProcess(p);
                if (package != null)
                {
                    Show(0, package);
                    Console.WriteLine();
                    Console.WriteLine();
                }
            }
        }*/

        /*public static void Show(int indent, AppxPackage package)
        {
            string sindent = new string(' ', indent);
            Console.WriteLine(sindent + "FullName               : " + package.FullName);
            Console.WriteLine(sindent + "FamilyName             : " + package.FamilyName);
            Console.WriteLine(sindent + "IsFramework            : " + package.IsFramework);
            Console.WriteLine(sindent + "ApplicationUserModelId : " + package.ApplicationUserModelId);
            Console.WriteLine(sindent + "Path                   : " + package.Path);
            Console.WriteLine(sindent + "Publisher              : " + package.Publisher);
            Console.WriteLine(sindent + "PublisherId            : " + package.PublisherId);
            Console.WriteLine(sindent + "Logo                   : " + package.Logo);
            Console.WriteLine(sindent + "Best Logo Path         : " + package.FindHighestScaleQualifiedImagePath(package.Logo));
            Console.WriteLine(sindent + "ProcessorArchitecture  : " + package.ProcessorArchitecture);
            Console.WriteLine(sindent + "Version                : " + package.Version);
            Console.WriteLine(sindent + "PublisherDisplayName   : " + package.PublisherDisplayName);
            Console.WriteLine(sindent + "   Localized           : " + package.LoadResourceString(package.PublisherDisplayName));
            Console.WriteLine(sindent + "DisplayName            : " + package.DisplayName);
            Console.WriteLine(sindent + "   Localized           : " + package.LoadResourceString(package.DisplayName));
            Console.WriteLine(sindent + "Description            : " + package.Description);
            Console.WriteLine(sindent + "   Localized           : " + package.LoadResourceString(package.Description));

            Console.WriteLine(sindent + "Apps                   :");
            int i = 0;
            foreach (var app in package.Apps)
            {
                Console.WriteLine(sindent + " App [" + i + "] Description       : " + app.Description);
                Console.WriteLine(sindent + "   Localized           : " + package.LoadResourceString(app.Description));
                Console.WriteLine(sindent + " App [" + i + "] DisplayName       : " + app.DisplayName);
                Console.WriteLine(sindent + "   Localized           : " + package.LoadResourceString(app.DisplayName));
                Console.WriteLine(sindent + " App [" + i + "] ShortName         : " + app.ShortName);
                Console.WriteLine(sindent + "   Localized           : " + package.LoadResourceString(app.ShortName));
                Console.WriteLine(sindent + " App [" + i + "] EntryPoint        : " + app.EntryPoint);
                Console.WriteLine(sindent + " App [" + i + "] Executable        : " + app.Executable);
                Console.WriteLine(sindent + " App [" + i + "] Id                : " + app.Id);
                Console.WriteLine(sindent + " App [" + i + "] Logo              : " + app.Logo);
                Console.WriteLine(sindent + " App [" + i + "] SmallLogo         : " + app.SmallLogo);
                Console.WriteLine(sindent + " App [" + i + "] StartPage         : " + app.StartPage);
                Console.WriteLine(sindent + " App [" + i + "] Square150x150Logo : " + app.Square150x150Logo);
                Console.WriteLine(sindent + " App [" + i + "] Square30x30Logo   : " + app.Square30x30Logo);
                Console.WriteLine(sindent + " App [" + i + "] BackgroundColor   : " + app.BackgroundColor);
                Console.WriteLine(sindent + " App [" + i + "] ForegroundText    : " + app.ForegroundText);
                Console.WriteLine(sindent + " App [" + i + "] WideLogo          : " + app.WideLogo);
                Console.WriteLine(sindent + " App [" + i + "] Wide310x310Logo   : " + app.Wide310x310Logo);
                Console.WriteLine(sindent + " App [" + i + "] Square310x310Logo : " + app.Square310x310Logo);
                Console.WriteLine(sindent + " App [" + i + "] Square70x70Logo   : " + app.Square70x70Logo);
                Console.WriteLine(sindent + " App [" + i + "] MinWidth          : " + app.MinWidth);
                Console.WriteLine(sindent + " App [" + i + "] Square71x71Logo   : " + app.GetStringValue("Square71x71Logzo"));
                i++;
            }

            Console.WriteLine(sindent + "Deps                   :");
            foreach (var dep in package.DependencyGraph)
            {
                Show(indent + 1, dep);
            }
        }

        public sealed class AppxPackage
        {
            public List<AppxApp> _apps = new List<AppxApp>();
            public IAppxManifestProperties _properties;

            public AppxPackage()
            {

            }

            public string FullName { get; private set; }
            public string Path { get; private set; }
            public string Publisher { get; private set; }
            public string PublisherId { get; private set; }
            public string ResourceId { get; private set; }
            public string FamilyName { get; private set; }
            public string ApplicationUserModelId { get; private set; }
            public string Logo { get; private set; }
            public string PublisherDisplayName { get; private set; }
            public string Description { get; private set; }
            public string DisplayName { get; private set; }
            public bool IsFramework { get; private set; }
            public Version Version { get; private set; }
            public AppxPackageArchitecture ProcessorArchitecture { get; private set; }

            public IReadOnlyList<AppxApp> Apps
            {
                get
                {
                    return _apps;
                }
            }

            public IEnumerable<AppxPackage> DependencyGraph
            {
                get
                {
                    return QueryPackageInfo(FullName, PackageConstants.PACKAGE_FILTER_ALL_LOADED).Where(p => p.FullName != FullName);
                }
            }

            public string FindHighestScaleQualifiedImagePath(string resourceName)
            {
                if (resourceName == null)
                    throw new ArgumentNullException("resourceName");

                const string scaleToken = ".scale-";
                var sizes = new List<int>();
                string name = System.IO.Path.GetFileNameWithoutExtension(resourceName);
                string ext = System.IO.Path.GetExtension(resourceName);
                foreach (var file in Directory.EnumerateFiles(System.IO.Path.Combine(Path, System.IO.Path.GetDirectoryName(resourceName)), name + scaleToken + "*" + ext))
                {
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                    int pos = fileName.IndexOf(scaleToken) + scaleToken.Length;
                    string sizeText = fileName.Substring(pos);
                    if (int.TryParse(sizeText, out int size))
                    {
                        sizes.Add(size);
                    }
                }
                if (sizes.Count == 0)
                    return null;

                sizes.Sort();
                return System.IO.Path.Combine(Path, System.IO.Path.GetDirectoryName(resourceName), name + scaleToken + sizes.Last() + ext);
            }

            public override string ToString()
            {
                return FullName;
            }

            public static AppxPackage FromWindow(IntPtr handle)
            {
                GetWindowThreadProcessId(handle, out int processId);
                if (processId == 0)
                    return null;

                return FromProcess(processId);
            }

            public static AppxPackage FromProcess(Process process)
            {
                if (process == null)
                {
                    process = Process.GetCurrentProcess();
                }

                try
                {
                    return FromProcess(process.Handle);
                }
                catch
                {
                    // probably access denied on .Handle
                    return null;
                }
            }

            public static AppxPackage FromProcess(int processId)
            {
                const int QueryLimitedInformation = 0x1000;
                IntPtr hProcess = OpenProcess(QueryLimitedInformation, false, processId);
                try
                {
                    return FromProcess(hProcess);
                }
                finally
                {
                    if (hProcess != IntPtr.Zero)
                    {
                        CloseHandle(hProcess);
                    }
                }
            }

            public static AppxPackage FromProcess(IntPtr hProcess)
            {
                if (hProcess == IntPtr.Zero)
                    return null;

                // hprocess must have been opened with QueryLimitedInformation
                int len = 0;
                GetPackageFullName(hProcess, ref len, null);
                if (len == 0)
                    return null;

                var sb = new StringBuilder(len);
                string fullName = GetPackageFullName(hProcess, ref len, sb) == 0 ? sb.ToString() : null;
                if (string.IsNullOrEmpty(fullName)) // not an AppX
                    return null;

                var package = QueryPackageInfo(fullName, PackageConstants.PACKAGE_FILTER_HEAD).First();

                len = 0;
                GetApplicationUserModelId(hProcess, ref len, null);
                sb = new StringBuilder(len);
                package.ApplicationUserModelId = GetApplicationUserModelId(hProcess, ref len, sb) == 0 ? sb.ToString() : null;
                return package;
            }

            public string GetPropertyStringValue(string name)
            {
                if (name == null)
                    throw new ArgumentNullException("name");

                return GetStringValue(_properties, name);
            }

            public bool GetPropertyBoolValue(string name)
            {
                if (name == null)
                    throw new ArgumentNullException("name");

                return GetBoolValue(_properties, name);
            }

            public string LoadResourceString(string resource)
            {
                return LoadResourceString(FullName, resource);
            }

            public static IEnumerable<AppxPackage> QueryPackageInfo(string fullName, PackageConstants flags)
            {
                OpenPackageInfoByFullName(fullName, 0, out IntPtr infoRef);
                if (infoRef != IntPtr.Zero)
                {
                    IntPtr infoBuffer = IntPtr.Zero;
                    try
                    {
                        int len = 0;
                        GetPackageInfo(infoRef, flags, ref len, IntPtr.Zero, out int count);
                        if (len > 0)
                        {
                            var factory = (IAppxFactory)new AppxFactory();
                            infoBuffer = Marshal.AllocHGlobal(len);
                            int res = GetPackageInfo(infoRef, flags, ref len, infoBuffer, out count);
                            for (int i = 0; i < count; i++)
                            {
                                var info = (PACKAGE_INFO)Marshal.PtrToStructure(infoBuffer + i * Marshal.SizeOf(typeof(PACKAGE_INFO)), typeof(PACKAGE_INFO));
                                var package = new AppxPackage
                                {
                                    FamilyName = Marshal.PtrToStringUni(info.packageFamilyName),
                                    FullName = Marshal.PtrToStringUni(info.packageFullName),
                                    Path = Marshal.PtrToStringUni(info.path),
                                    Publisher = Marshal.PtrToStringUni(info.packageId.publisher),
                                    PublisherId = Marshal.PtrToStringUni(info.packageId.publisherId),
                                    ResourceId = Marshal.PtrToStringUni(info.packageId.resourceId),
                                    ProcessorArchitecture = info.packageId.processorArchitecture,
                                    Version = new Version(info.packageId.VersionMajor, info.packageId.VersionMinor, info.packageId.VersionBuild, info.packageId.VersionRevision)
                                };

                                // read manifest
                                string manifestPath = System.IO.Path.Combine(package.Path, "AppXManifest.xml");
                                const int STGM_SHARE_DENY_NONE = 0x40;
                                SHCreateStreamOnFileEx(manifestPath, STGM_SHARE_DENY_NONE, 0, false, IntPtr.Zero, out IStream strm);
                                if (strm != null)
                                {
                                    var reader = factory.CreateManifestReader(strm);
                                    package._properties = reader.GetProperties();
                                    package.Description = package.GetPropertyStringValue("Description");
                                    package.DisplayName = package.GetPropertyStringValue("DisplayName");
                                    package.Logo = package.GetPropertyStringValue("Logo");
                                    package.PublisherDisplayName = package.GetPropertyStringValue("PublisherDisplayName");
                                    package.IsFramework = package.GetPropertyBoolValue("Framework");

                                    var apps = reader.GetApplications();
                                    while (apps.GetHasCurrent())
                                    {
                                        var app = apps.GetCurrent();
                                        var appx = new AppxApp(app)
                                        {
                                            Description = GetStringValue(app, "Description"),
                                            DisplayName = GetStringValue(app, "DisplayName"),
                                            EntryPoint = GetStringValue(app, "EntryPoint"),
                                            Executable = GetStringValue(app, "Executable"),
                                            Id = GetStringValue(app, "Id"),
                                            Logo = GetStringValue(app, "Logo"),
                                            SmallLogo = GetStringValue(app, "SmallLogo"),
                                            StartPage = GetStringValue(app, "StartPage"),
                                            Square150x150Logo = GetStringValue(app, "Square150x150Logo"),
                                            Square30x30Logo = GetStringValue(app, "Square30x30Logo"),
                                            BackgroundColor = GetStringValue(app, "BackgroundColor"),
                                            ForegroundText = GetStringValue(app, "ForegroundText"),
                                            WideLogo = GetStringValue(app, "WideLogo"),
                                            Wide310x310Logo = GetStringValue(app, "Wide310x310Logo"),
                                            ShortName = GetStringValue(app, "ShortName"),
                                            Square310x310Logo = GetStringValue(app, "Square310x310Logo"),
                                            Square70x70Logo = GetStringValue(app, "Square70x70Logo"),
                                            MinWidth = GetStringValue(app, "MinWidth")
                                        };
                                        package._apps.Add(appx);
                                        apps.MoveNext();
                                    }
                                    Marshal.ReleaseComObject(strm);
                                }
                                yield return package;
                            }
                            Marshal.ReleaseComObject(factory);
                        }
                    }
                    finally
                    {
                        if (infoBuffer != IntPtr.Zero)
                        {
                            Marshal.FreeHGlobal(infoBuffer);
                        }
                        ClosePackageInfo(infoRef);
                    }
                }
            }

            public static string LoadResourceString(string packageFullName, string resource)
            {
                if (packageFullName == null)
                    throw new ArgumentNullException("packageFullName");

                if (string.IsNullOrWhiteSpace(resource))
                    return null;

                const string resourceScheme = "ms-resource:";
                if (!resource.StartsWith(resourceScheme))
                    return null;

                string part = resource.Substring(resourceScheme.Length);
                string url;

                if (part.StartsWith("/"))
                {
                    url = resourceScheme + "//" + part;
                }
                else
                {
                    url = resourceScheme + "///resources/" + part;
                }

                string source = string.Format("@{{{0}? {1}}}", packageFullName, url);
                var sb = new StringBuilder(1024);
                int i = SHLoadIndirectString(source, sb, sb.Capacity, IntPtr.Zero);
                if (i != 0)
                    return null;

                return sb.ToString();
            }

            public static string GetStringValue(IAppxManifestProperties props, string name)
            {
                if (props == null)
                    return null;

                props.GetStringValue(name, out string value);
                return value;
            }

            public static bool GetBoolValue(IAppxManifestProperties props, string name)
            {
                props.GetBoolValue(name, out bool value);
                return value;
            }

            internal static string GetStringValue(IAppxManifestApplication app, string name)
            {
                app.GetStringValue(name, out string value);
                return value;
            }

            [Guid("5842a140-ff9f-4166-8f5c-62f5b7b0c781"), ComImport]
            public class AppxFactory
            {
            }

            [Guid("BEB94909-E451-438B-B5A7-D79E767B75D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IAppxFactory
            {
                void _VtblGap0_2(); // skip 2 methods
                IAppxManifestReader CreateManifestReader(IStream inputStream);
            }

            [Guid("4E1BD148-55A0-4480-A3D1-15544710637C"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IAppxManifestReader
            {
                void _VtblGap0_1(); // skip 1 method
                IAppxManifestProperties GetProperties();
                void _VtblGap1_5(); // skip 5 methods
                IAppxManifestApplicationsEnumerator GetApplications();
            }

            [Guid("9EB8A55A-F04B-4D0D-808D-686185D4847A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IAppxManifestApplicationsEnumerator
            {
                IAppxManifestApplication GetCurrent();
                bool GetHasCurrent();
                bool MoveNext();
            }

            [Guid("5DA89BF4-3773-46BE-B650-7E744863B7E8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IAppxManifestApplication
            {
                [PreserveSig]
                int GetStringValue([MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.LPWStr)] out string vaue);
            }

            [Guid("03FAF64D-F26F-4B2C-AAF7-8FE7789B8BCA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IAppxManifestProperties
            {
                [PreserveSig]
                int GetBoolValue([MarshalAs(UnmanagedType.LPWStr)]string name, out bool value);
                [PreserveSig]
                int GetStringValue([MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.LPWStr)] out string vaue);
            }

            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
            public static extern int SHLoadIndirectString(string pszSource, StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved);

            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
            public static extern int SHCreateStreamOnFileEx(string fileName, int grfMode, int attributes, bool create, IntPtr reserved, out IStream stream);

            [DllImport("user32.dll")]
            public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

            [DllImport("kernel32.dll")]
            public static extern IntPtr OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);

            [DllImport("kernel32.dll")]
            public static extern bool CloseHandle(IntPtr hObject);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            public static extern int OpenPackageInfoByFullName(string packageFullName, int reserved, out IntPtr packageInfoReference);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            public static extern int GetPackageInfo(IntPtr packageInfoReference, PackageConstants flags, ref int bufferLength, IntPtr buffer, out int count);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            public static extern int ClosePackageInfo(IntPtr packageInfoReference);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            public static extern int GetPackageFullName(IntPtr hProcess, ref int packageFullNameLength, StringBuilder packageFullName);*/

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            public static extern int GetApplicationUserModelId(IntPtr hProcess, ref int applicationUserModelIdLength, StringBuilder applicationUserModelId);

            /*[Flags]
            public enum PackageConstants
            {
                PACKAGE_FILTER_ALL_LOADED = 0x00000000,
                PACKAGE_PROPERTY_FRAMEWORK = 0x00000001,
                PACKAGE_PROPERTY_RESOURCE = 0x00000002,
                PACKAGE_PROPERTY_BUNDLE = 0x00000004,
                PACKAGE_FILTER_HEAD = 0x00000010,
                PACKAGE_FILTER_DIRECT = 0x00000020,
                PACKAGE_FILTER_RESOURCE = 0x00000040,
                PACKAGE_FILTER_BUNDLE = 0x00000080,
                PACKAGE_INFORMATION_BASIC = 0x00000000,
                PACKAGE_INFORMATION_FULL = 0x00000100,
                PACKAGE_PROPERTY_DEVELOPMENT_MODE = 0x00010000,
            }

            [StructLayout(LayoutKind.Sequential, Pack = 4)]
            public struct PACKAGE_INFO
            {
                public int reserved;
                public int flags;
                public IntPtr path;
                public IntPtr packageFullName;
                public IntPtr packageFamilyName;
                public PACKAGE_ID packageId;
            }

            [StructLayout(LayoutKind.Sequential, Pack = 4)]
            public struct PACKAGE_ID
            {
                public int reserved;
                public AppxPackageArchitecture processorArchitecture;
                public ushort VersionRevision;
                public ushort VersionBuild;
                public ushort VersionMinor;
                public ushort VersionMajor;
                public IntPtr name;
                public IntPtr publisher;
                public IntPtr resourceId;
                public IntPtr publisherId;
            }
        }

        public class AppxApp
        {
            public AppxPackage.IAppxManifestApplication _app;

            internal AppxApp(AppxPackage.IAppxManifestApplication app)
            {
                _app = app;
            }

            public string GetStringValue(string name)
            {
                if (name == null)
                    throw new ArgumentNullException("name");

                return AppxPackage.GetStringValue(_app, name);
            }

            // we code well-known but there are others (like Square71x71Logo, Square44x44Logo, whatever ...)
            // https://msdn.microsoft.com/en-us/library/windows/desktop/hh446703.aspx
            public string Description { get; internal set; }
            public string DisplayName { get; internal set; }
            public string EntryPoint { get; internal set; }
            public string Executable { get; internal set; }
            public string Id { get; internal set; }
            public string Logo { get; internal set; }
            public string SmallLogo { get; internal set; }
            public string StartPage { get; internal set; }
            public string Square150x150Logo { get; internal set; }
            public string Square30x30Logo { get; internal set; }
            public string BackgroundColor { get; internal set; }
            public string ForegroundText { get; internal set; }
            public string WideLogo { get; internal set; }
            public string Wide310x310Logo { get; internal set; }
            public string ShortName { get; internal set; }
            public string Square310x310Logo { get; internal set; }
            public string Square70x70Logo { get; internal set; }
            public string MinWidth { get; internal set; }
        }

        public enum AppxPackageArchitecture
        {
            x86 = 0,
            Arm = 5,
            x64 = 9,
            Neutral = 11,
            Arm64 = 12
        }*/
    }
}