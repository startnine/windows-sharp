using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

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

                NativeMethods.GetWindowThreadProcessId(hWnd, out uint windowPID);

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
    }
}
