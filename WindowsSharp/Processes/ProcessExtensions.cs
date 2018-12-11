using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace WindowsSharp.Processes
{
    public static class ProcessExtensions
    {
        public static string GetExecutablePath(this Process process)
        {
            string returnValue = string.Empty;
            StringBuilder stringBuilder = new StringBuilder(1024);
            IntPtr hprocess = NativeMethods.OpenProcess(0x1000, false, process.Id);

            if (hprocess != IntPtr.Zero)
            {
                /*try
                {*/
                    int size = stringBuilder.Capacity;

                    if (NativeMethods.QueryFullProcessImageName(hprocess, 0, stringBuilder, out size))
                        returnValue = stringBuilder.ToString();
                /*}
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
                }*/
            }
            /*var package = AppxPackage.FromProcess(process);
            if (package != null)
            {
                Debug.WriteLine("PACKAGE.APPLICATIONUSERMODELID: " + package.ApplicationUserModelId);
                returnValue = package.ApplicationUserModelId;
            }*/
            /*else
                Debug.WriteLine("PACKAGE IS NULL");*/

            //Debug.WriteLine("returnValue: " + returnValue);
            return returnValue;
        }

        /*static ProcessExtensions()
        {
            var form = new ProcessWindowMonitorForm();

            while (true)
                System.Windows.Forms.Application.DoEvents();
        }*/


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

        public static string GetProcessAppUserModelId(this Process process)
        {
            string output = string.Empty;

            IntPtr handle = process.Handle;

            int outputLength = 0;
            AppxMethods.GetApplicationUserModelId(handle, ref outputLength, null);

            StringBuilder outputBuilder = new StringBuilder(outputLength);


            int resultValue = 1; AppxMethods.GetApplicationUserModelId(handle, ref outputLength, outputBuilder);

            if (resultValue == 0)
                output = outputBuilder.ToString();

            return output;
        }

        public static bool IsSameApp(this Process process, Process compareProcess)
        {
            bool pathsEqual = process.GetExecutablePath().ToLowerInvariant() == compareProcess.GetExecutablePath().ToLowerInvariant();
            bool appUserModelIdsEqual = true;

            if (Environment.OSVersion.Version >= new Version(6, 2, 8400, 0))
                appUserModelIdsEqual = process.GetProcessAppUserModelId() == compareProcess.GetProcessAppUserModelId();

            return pathsEqual && appUserModelIdsEqual;
        }

        public static bool BelongsToExecutable(this Process process, DiskItems.DiskItem compareItem)
        {
            return process.GetExecutablePath().ToLowerInvariant() == compareItem.ItemPath.ToLowerInvariant();
        }
    }
}