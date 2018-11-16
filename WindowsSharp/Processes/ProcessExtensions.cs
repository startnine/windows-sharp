using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Diagnostics;



using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Runtime.ConstrainedExecution;
using System.Security;
using System.Runtime.CompilerServices;

namespace WindowsSharp.Processes
{
    public static class ProcessExtensions
    {
        static ProcessExtensions()
        {
            var form = new ProcessWindowMonitorForm();

            while (true)
                System.Windows.Forms.Application.DoEvents();
        }


        public static List<ProcessWindow> Windows(this Process process)
        {
            List<ProcessWindow> windows = new List<ProcessWindow>();

            List<IntPtr> collection = new List<IntPtr>();

            int processPID = process.Id;
            Boolean Filter(IntPtr hWnd, Int32 lParam)
            {
                var strbTitle = new StringBuilder(NativeMethods.GetWindowTextLength(hWnd));
                NativeMethods.GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
                var strTitle = strbTitle.ToString();

                uint windowPID = 0;
                NativeMethods.GetWindowThreadProcessId(hWnd, out windowPID);

                if ((NativeMethods.IsWindowVisible(hWnd) && string.IsNullOrEmpty(strTitle) == false) && (windowPID == processPID))
                    collection.Add(hWnd);

                return true;
            }

            if (NativeMethods.EnumDesktopWindows(IntPtr.Zero, Filter, IntPtr.Zero))
            {
                foreach (var hwnd in collection)
                    windows.Add(new ProcessWindow(hwnd));
            }

            return windows;
        }

        public static string GetExecutablePath(Process process)
        {
            string returnValue = "";
            StringBuilder stringBuilder = new StringBuilder(1024);
            IntPtr hprocess = NativeMethods.OpenProcess(0x1000, false, process.Id);

            if (hprocess != IntPtr.Zero)
            {
                try
                {
                    int size = stringBuilder.Capacity;

                    if (NativeMethods.QueryFullProcessImageName(hprocess, 0, stringBuilder, out size))
                        returnValue = stringBuilder.ToString();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                    try
                    {
                        returnValue = process.MainModule.FileName;
                    }
                    catch (Exception exc)
                    {
                        Debug.WriteLine(exc);
                    }
                }
            }

            return returnValue;
        }

        public static bool IsSameApp(this Process process, Process compareProcess)
        {
            return (GetExecutablePath(process).ToLowerInvariant()) == (GetExecutablePath(compareProcess).ToLowerInvariant());
        }

        public static bool IsSameApp(this Process process, DiskItems.DiskItem compareItem)
        {
            return (GetExecutablePath(process).ToLowerInvariant()) == (compareItem.ItemPath.ToLowerInvariant());
        }

        public static class NativeMethods
        {
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
            public static extern int TrackPopupMenuEx(IntPtr hmenu, uint fuFlags,
              int x, int y, IntPtr hwnd, IntPtr lptpm);

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

            [DllImport("user32.dll")]
            public static extern Boolean SetForegroundWindow(IntPtr hWnd);

            [DllImport("User32")]
            public static extern Int32 ShowWindow(IntPtr hwnd, Int32 nCmdShow);

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
            public const Int32 GwlExStyle = -20;
            public const Int32 WsExToolWindow = 0x00000080;

            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();
        }
    }
}
