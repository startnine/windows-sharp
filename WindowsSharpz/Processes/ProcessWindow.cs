using System;
using System.Collections.Generic;
using System.Text;
//using System.Windows.Forms;
using static WindowsSharp.Processes.ProcessExtensions;

namespace WindowsSharp.Processes
{
    public class ProcessWindow
    {
        public IntPtr handle
        {
            get;
            private set;
        }

        public string Title
        {
            get
            {
                var strbTitle = new StringBuilder(NativeMethods.GetWindowTextLength(handle));
                NativeMethods.GetWindowText(handle, strbTitle, strbTitle.Capacity + 1);
                return strbTitle.ToString();
            }
        }

        public bool Close()
        {
            return NativeMethods.PostMessage(handle, NativeMethods.WmClose, IntPtr.Zero, IntPtr.Zero);
        }

        public event EventHandler<EventArgs> Closed;

        public Int32 Maximize()
        {
            return NativeMethods.ShowWindow(handle, NativeMethods.SwMaximize);
        }

        public Int32 Minimize()
        {
            return NativeMethods.ShowWindow(handle, NativeMethods.SwForceMinimize);
        }

        public Int32 Restore()
        {
            return NativeMethods.ShowWindow(handle, NativeMethods.SwRestore);
        }

        public ProcessWindow(IntPtr windowHandle)
        {
            handle = windowHandle;
            System.Timers.Timer timer = new System.Timers.Timer(1);
            timer.Elapsed += (sneder, args) =>
            {
                if (!NativeMethods.IsWindow(handle))
                {
                    Closed.Invoke(this, new EventArgs());
                    timer.Stop();
                }
            };
            timer.Start();
        }
    }
}
