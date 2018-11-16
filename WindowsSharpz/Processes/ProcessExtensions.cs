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

namespace WindowsSharp.Processes
{
    public static class ProcessExtensions
    {
        static ProcessExtensions()
        {

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


        public static class NativeMethods
        {
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
        }
    }
}
