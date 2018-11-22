using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
//using System.Windows.Forms;
using System.Windows.Automation;
using System.Windows.Forms;
using Timer = System.Timers.Timer;
using static WindowsSharp.AppxMethods;

namespace WindowsSharp.Processes
{
    public class ProcessWindow : INotifyPropertyChanged
    {
        private static ProcessWindow _activeWindow = new ProcessWindow(NativeMethods.GetForegroundWindow());
        public static ProcessWindow ActiveWindow
        {
            get => _activeWindow;
            internal set
            {
                _activeWindow = value;
                ActiveWindowChanged?.Invoke(null, new WindowEventArgs(_activeWindow));
            }
        }

        public static event EventHandler<WindowEventArgs> WindowOpened;

        internal static void RaiseWindowOpened(IntPtr hwnd)
        {
            RaiseWindowOpened(hwnd, false);
        }

        internal static void RaiseWindowOpened(IntPtr hwnd, bool isVisibilityChange)
        {
            ProcessWindow.WindowOpened?.Invoke(null, new WindowEventArgs(new ProcessWindow(hwnd))
            {
                IsVisibilityChange = isVisibilityChange
            });
        }

        public static event EventHandler<WindowEventArgs> ActiveWindowChanged;

        public static event EventHandler<WindowEventArgs> WindowClosed;

        private void NotifyPropertyChanged(string info)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(info));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        internal static void RaiseWindowClosed(IntPtr hwnd)
        {
            RaiseWindowOpened(hwnd, false);
        }

        internal static void RaiseWindowClosed(IntPtr hwnd, bool isVisibilityChange)
        {
            ProcessWindow.WindowClosed?.Invoke(null, new WindowEventArgs(new ProcessWindow(hwnd))
            {
                IsVisibilityChange = isVisibilityChange
            });
            //Debug.WriteLine("WindowClosed");
        }

        /*static IEnumerable<ProcessWindow> GetMetroApps()
        {
            if ((Environment.OSVersion.Version >= new Version(6, 2, 8400, 0)) && (Environment.OSVersion.Version < new Version(10, 0, 10240, 0)))
            {
                foreach (var p in Process.GetProcesses())
                {
                    AppxPackage package = AppxPackage.FromProcess(p);
                    if (package != null)
                    {
                        yield return new ProcessWindow(package);
                    }
                }
            }
            else
            {
                yield break;
                //return new List<ProcessWindow>();
            }
        }*/

        public static IEnumerable<ProcessWindow> RealProcessWindows
        {
            get
            {
                var collection = new List<IntPtr>();

                Boolean Filter(IntPtr hWnd, Int32 lParam)
                {
                    var strbTitle = new StringBuilder(NativeMethods.GetWindowTextLength(hWnd));
                    NativeMethods.GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
                    var strTitle = strbTitle.ToString();


                    if (NativeMethods.IsWindowVisible(hWnd) && string.IsNullOrEmpty(strTitle) == false)
                        collection.Add(hWnd);

                    return true;
                }

                if (!NativeMethods.EnumDesktopWindows(IntPtr.Zero, Filter, IntPtr.Zero)) yield break;

                foreach (var hwnd in collection)
                    yield return new ProcessWindow(hwnd);

                /*foreach (var w in GetMetroApps())
                    yield return w;*/
            }
        }

        //static Int32 TASKSTYLE = 0x10000000;// | 0x00800000;

        public static IEnumerable<ProcessWindow> ProcessWindows
        {
            /*get
            {
                var collection = new List<IntPtr>();

                Boolean Filter(IntPtr hWnd, Int32 lParam)
                {
                    var strbTitle = new StringBuilder(NativeMethods.GetWindowTextLength(hWnd));
                    NativeMethods.GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
                    var strTitle = strbTitle.ToString();


                    if (NativeMethods.IsWindowVisible(hWnd) && string.IsNullOrEmpty(strTitle) == false)
                        collection.Add(hWnd);

                    return true;
                }

                if (!NativeMethods.EnumDesktopWindows(IntPtr.Zero, Filter, IntPtr.Zero)) yield break;

                foreach (var hwnd in collection)
                {
                    try
                    {
                        if (IsWindowUserAccessible(hwnd))
                        {
                            yield return new ProcessWindow(hwnd);
                        }
                    }
                    finally
                    {

                    }
                }
            }*/
            get
            {
                var collection = new List<IntPtr>();

                Boolean Filter(IntPtr hWnd, Int32 lParam)
                {
                    var strbTitle = new StringBuilder(NativeMethods.GetWindowTextLength(hWnd));
                    NativeMethods.GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
                    var strTitle = strbTitle.ToString();


                    if (NativeMethods.IsWindowVisible(hWnd) && string.IsNullOrEmpty(strTitle) == false)
                        collection.Add(hWnd);

                    return true;
                }

                if (!NativeMethods.EnumDesktopWindows(IntPtr.Zero, Filter, IntPtr.Zero)) yield break;

                foreach (var hwnd in collection)
                {
                    try
                    {
                        /*if (Environment.Is64BitProcess)
                        {
                            if ((TASKSTYLE == (TASKSTYLE & NativeMethods.GetWindowLong(hwnd, NativeMethods.GwlStyle).ToInt64())) && ((NativeMethods.GetWindowLong(hwnd, NativeMethods.GwlExStyle).ToInt64() & NativeMethods.WsExToolWindow) != NativeMethods.WsExToolWindow))
                            {
                                yield return new ProcessWindow(hwnd);
                            }
                        }
                        else
                        {
                            if ((TASKSTYLE == (TASKSTYLE & NativeMethods.GetWindowLong(hwnd, NativeMethods.GwlStyle).ToInt32())) && ((NativeMethods.GetWindowLong(hwnd, NativeMethods.GwlExStyle).ToInt32() & NativeMethods.WsExToolWindow) != NativeMethods.WsExToolWindow))
                            {
                                yield return new ProcessWindow(hwnd);
                            }
                        }*/
                        if (IsWindowUserAccessible(hwnd))
                            yield return new ProcessWindow(hwnd);
                    }
                    finally
                    {

                    }
                }

                /*foreach (var w in GetMetroApps())
                    yield return w;*/
            }
        }

        public static bool IsWindowUserAccessible(IntPtr hwnd)
        {
            var style = NativeMethods.GetWindowLong(hwnd, NativeMethods.GwlStyle);
            var exStyle = NativeMethods.GetWindowLong(hwnd, NativeMethods.GwlExStyle);
            var exStyleToolWindow = exStyle;
            //exStyleToolWindow |= NativeMethods.WsExToolWindow;

            if (Environment.Is64BitProcess)
                return //(NativeMethods.WsVisible == (style.ToInt64() & NativeMethods.WsVisible))
                       //&& ((style.ToInt64() & NativeMethods.WsChild) == style.ToInt64())
                    NativeMethods.IsWindowVisible(hwnd)
                    //&& (NativeMethods.WsBorder == (style.ToInt64() & NativeMethods.WsBorder))
                    && ((style.ToInt64() & ~NativeMethods.WsChild) == style.ToInt64())
                    && ((exStyle.ToInt64() & ~NativeMethods.WsExToolWindow) == exStyle.ToInt64());
            else
                return //(NativeMethods.WsVisible == (style.ToInt32() & NativeMethods.WsVisible))
                       //&& ((style.ToInt32() & NativeMethods.WsChild) == style.ToInt32())
                    NativeMethods.IsWindowVisible(hwnd)
                    //&& (NativeMethods.WsBorder == (style.ToInt32() & NativeMethods.WsBorder))
                    && ((style.ToInt32() & ~NativeMethods.WsChild) == style.ToInt32())
                    && ((exStyle.ToInt32() & ~NativeMethods.WsExToolWindow) == exStyle.ToInt32());
        }

        static ProcessWindow()
        {
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Children,
                (sender, e) => WindowOpened?.Invoke(null,
                                                    new WindowEventArgs(
                                                        new ProcessWindow(new IntPtr(((AutomationElement)sender).Current.NativeWindowHandle))
                                                        )));

            /*Automation.AddAutomationEventHandler(
				WindowPattern.WindowClosedEvent,
				AutomationElement.RootElement,
				TreeScope.Subtree,
				(sender, e) => WindowClosed?.Invoke(null,
													new WindowEventArgs(
														new ProcessWindow(new IntPtr(((AutomationElement) sender).Cached.NativeWindowHandle)))));*/

            Timer timer = new Timer(1);
            timer.Elapsed += (sneder, args) =>
            {
                //Debug.WriteLine("timer elapsed");
                if (ActiveWindow != null)
                {
                    //Debug.WriteLine("active window is non-null");
                    IntPtr newActive = NativeMethods.GetForegroundWindow();

                    if ((ActiveWindow.Handle != newActive) && IsWindowUserAccessible(newActive))
                        ActiveWindow = new ProcessWindow(newActive);
                }
            };
            timer.Start();
        }

        public IntPtr Handle
        {
            get;
            private set;
        }

        /*public AppxPackage Package
        {
            get;
            private set;
        }*/

        public Process Process
        {
            get
            {
                NativeMethods.GetWindowThreadProcessId(Handle, out var pid);
                return Process.GetProcessById((Int32)pid);
            }
        }

        string _windowTitle = string.Empty;

        public string Title
        {
            get
            {
                if (_windowTitle == string.Empty)
                    _windowTitle = GetWindowTitle();
                /*var strbTitle = new StringBuilder(NativeMethods.GetWindowTextLength(Handle));
                NativeMethods.GetWindowText(Handle, strbTitle, strbTitle.Capacity + 1);
                return strbTitle.ToString();*/

                return _windowTitle;
            }
            internal set
            {
                /*try
                {*/
                    _windowTitle = value;
                /*}
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }*/
                NotifyPropertyChanged("Title");
            }
        }

        string GetWindowTitle()
        {
            var strbTitle = new StringBuilder(NativeMethods.GetWindowTextLength(Handle));
            NativeMethods.GetWindowText(Handle, strbTitle, strbTitle.Capacity + 1);
            string valueString = strbTitle.ToString();
            /*if ((Environment.OSVersion.Version >= new Version(6, 2, 8400, 0)) && string.IsNullOrWhiteSpace(valueString) && (Package != null))
                valueString = Package.DisplayName;*/

            return valueString;
        }

        //Rectangle _windowRectangle = new Rectangle();

        Rectangle GetBounds()
        {
            Rectangle rectangle = new Rectangle();

            if (NativeMethods.GetWindowRect(Handle, out NativeMethods.RECT rect))
                rectangle = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);

            return rectangle;
        }

        public Rectangle WindowBounds
        {
            get => GetBounds();
            set
            {
                Rectangle rect = value;
                NativeMethods.MoveWindow(Handle, value.X, value.Y, value.Width, value.Height, true);
                NotifyPropertyChanged("WindowBounds");
            }
        }

        /// <summary>
        /// Gets the icon of the window, preferring the large variant.
        /// </summary>
		public Icon Icon
        {
            get
            {
                /*if ((Environment.OSVersion.Version >= new Version(6, 2, 8400, 0)) && (Package != null))
                {
                    string path = Environment.ExpandEnvironmentVariables(Package.Logo);
                    if (System.IO.File.Exists(path))
                        return new Icon(Package.Logo);
                    else
                        return null;
                }
                else
                {*/
                    var iconHandle = NativeMethods.SendMessage(Handle, NativeMethods.WmGetIcon, NativeMethods.IconBig, 0);

                    if (iconHandle == IntPtr.Zero)
                        iconHandle = NativeMethods.GetClassLongPtr(Handle, NativeMethods.GclHIcon);
                    if (iconHandle == IntPtr.Zero)
                        iconHandle = NativeMethods.SendMessage(Handle, NativeMethods.WmGetIcon, NativeMethods.IconSmall, 0);
                    if (iconHandle == IntPtr.Zero)
                        iconHandle = NativeMethods.SendMessage(Handle, NativeMethods.WmGetIcon, NativeMethods.IconSmall2, 0);
                    if (iconHandle == IntPtr.Zero)
                        iconHandle = NativeMethods.GetClassLongPtr(Handle, NativeMethods.GclHIconSm);

                    if (iconHandle == IntPtr.Zero)
                        return null;

                    try
                    {
                        return Icon.FromHandle(iconHandle);
                    }
                    finally
                    {
                        NativeMethods.DestroyIcon(iconHandle);
                    }
                //}
            }
        }

        /// <summary>
        /// Gets the icon of the window, preferring the large variant.
        /// </summary>
		public Icon SmallIcon
        {
            get
            {
                var iconHandle = NativeMethods.SendMessage(Handle, NativeMethods.WmGetIcon, NativeMethods.IconSmall2, 0);

                if (iconHandle == IntPtr.Zero)
                    iconHandle = NativeMethods.SendMessage(Handle, NativeMethods.WmGetIcon, NativeMethods.IconSmall, 0);
                if (iconHandle == IntPtr.Zero)
                    iconHandle = NativeMethods.SendMessage(Handle, NativeMethods.WmGetIcon, NativeMethods.IconBig, 0);
                if (iconHandle == IntPtr.Zero)
                    iconHandle = NativeMethods.GetClassLongPtr(Handle, NativeMethods.GclHIcon);
                if (iconHandle == IntPtr.Zero)
                    iconHandle = NativeMethods.GetClassLongPtr(Handle, NativeMethods.GclHIconSm);

                if (iconHandle == IntPtr.Zero)
                    return null;

                try
                {
                    return Icon.FromHandle(iconHandle);
                }
                finally
                {
                    NativeMethods.DestroyIcon(iconHandle);
                }
            }
        }

        public bool IsMinimized
        {
            get
            {
                var placement = GetPlacement();
                return (placement.ShowCmd == NativeMethods.SwShowMinimized);
            }
        }

        public bool IsMaximized
        {
            get
            {
                var placement = GetPlacement();
                return (placement.ShowCmd == NativeMethods.SwShowMaximized);
            }
        }

        bool _isVisible;
        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                _isVisible = value;
                NotifyPropertyChanged("IsVisible");
            }
        }

        static Image CaptureWindow(IntPtr handle)
        {
            try
            {
                IntPtr hdcSrc = NativeMethods.GetWindowDC(handle);

                NativeMethods.RECT windowRect = new NativeMethods.RECT();
                NativeMethods.GetWindowRect(handle, out windowRect);

                int width = windowRect.Right - windowRect.Left;
                int height = windowRect.Bottom - windowRect.Top;

                IntPtr hdcDest = NativeMethods.CreateCompatibleDC(hdcSrc);
                IntPtr hBitmap = NativeMethods.CreateCompatibleBitmap(hdcSrc, width, height);

                IntPtr hOld = NativeMethods.SelectObject(hdcDest, hBitmap);

                NativeMethods.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, 13369376);
                NativeMethods.SelectObject(hdcDest, hOld);
                NativeMethods.DeleteDC(hdcDest);
                NativeMethods.ReleaseDC(handle, hdcSrc);

                Image image = Image.FromHbitmap(hBitmap);
                NativeMethods.DeleteObject(hBitmap);

                NativeMethods.DeleteObject(hdcDest);

                return image;
            }
            catch (ExternalException ex)
            {
                Debug.WriteLine("ProcessWindow.CaptureWindow machine broke:\n" + ex);
                return null;
            }
        }

        Image GetThumbnail()
        {
            if (NativeMethods.DwmIsCompositionEnabled())
                return CaptureWindow(Handle);
            else
                return null;
        }

        //Image _thumbnail = null;
        public Image Thumbnail
        {
            get => GetThumbnail();
            private set
            {
                //_thumbnail = value;
                NotifyPropertyChanged("Thumbnail");
            }
        }

        public ProcessWindow(IntPtr windowHandle)
        {
            Handle = windowHandle;
            IsVisible = NativeMethods.IsWindowVisible(Handle);
            Thumbnail = GetThumbnail();

            Timer timer = new Timer(10);
            //{ Interval = 10 };

            timer.Elapsed += (sneder, args) =>
            {
                if (!NativeMethods.IsWindow(Handle))
                {
                    try
                    {
                        ProcessWindow.WindowClosed?.Invoke(null, new WindowEventArgs(this));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }
                    timer.Stop();
                }
                else
                {
                    string currentTitle = GetWindowTitle();
                    if (Title != currentTitle)
                        Title = currentTitle;

                    bool visibleNow = NativeMethods.IsWindowVisible(Handle);
                    /*(IsVisible && (!visibleNow))
                    || ((!IsVisible) && visibleNow)
                    )*/
                    if (visibleNow != IsVisible)
                    {
                        //Debug.WriteLine("IsVisible: " + IsVisible.ToString() + "\nvisibleNow: " + visibleNow.ToString());

                        //visibleNow;

                        if (visibleNow)
                        {
                            IsVisible = true;
                            RaiseWindowOpened(Handle, true);
                        }
                        else
                        {
                            IsVisible = false;
                            RaiseWindowClosed(Handle, true);
                        }
                    }

                    if (!IsMinimized)
                        NotifyPropertyChanged("WindowBounds");

                    if (NativeMethods.DwmIsCompositionEnabled())
                        NotifyPropertyChanged("Thumbnail");
                }
            };
            //SetWinEventHook
            timer.Start();

            Timer boundsTimer = new Timer(100);
            //{ Interval = 100 };

            boundsTimer.Elapsed += (sneder, args) =>
            {
                if (!NativeMethods.IsWindow(Handle))
                    boundsTimer.Stop();
                else
                {
                    //WindowBounds = GetBounds();
                    NotifyPropertyChanged("WindowBounds");

                    if (NativeMethods.DwmIsCompositionEnabled())
                        NotifyPropertyChanged("Thumbnail");
                    //Thumbnail = GetThumbnail();
                }
            };

            //boundsTimer.Start();
        }

        /*public ProcessWindow(AppxPackage package)
        {
            Package = package;
        }*/

        public bool Close()
        {
            if (Handle != null)
                return NativeMethods.PostMessage(Handle, NativeMethods.WmClose, IntPtr.Zero, IntPtr.Zero);
            else
                return false; //app
        }

        public bool Maximize()
        {
            //Debug.WriteLine("Maximize()");
            return NativeMethods.ShowWindowAsync(Handle, NativeMethods.SwMaximize);
        }

        public bool Minimize()
        {
            //Debug.WriteLine("Minimize()");
            return NativeMethods.ShowWindowAsync(Handle, NativeMethods.SwMinimize);
        }
         
        public bool Restore()
        {
            //Debug.WriteLine("Restore()");
            return NativeMethods.ShowWindowAsync(Handle, NativeMethods.SwRestore);
        }

        IntPtr _topmostHandle = IntPtr.Zero;

        public void Peek()
        {
            Peek(IntPtr.Zero);
        }

        public void Peek(IntPtr topmostHandle)
        {
            uint peekOn = 1;
            uint peekType = 1;
            _topmostHandle = topmostHandle;
            //IntPtr topmostHandle = IntPtr.Zero;
            //Debug.WriteLine("Peek target title: " + new ProcessWindow(Handle).Title);

            if (Environment.OSVersion.Version >= new Version(6, 2, 8400, 0))
                NativeMethods.DwmpActivateLivePreview(peekOn, Handle, topmostHandle, peekType, UIntPtr.Zero);
            else
                NativeMethods.DwmpActivateLivePreview(peekOn, Handle, topmostHandle, peekType);
        }

        public void Unpeek()
        {
            uint peekOn = 0;
            uint peekType = 1;
            //IntPtr topmostHandle = IntPtr.Zero;
            //Debug.WriteLine("Unpeek target title: " + new ProcessWindow(Handle).Title);

            if (Environment.OSVersion.Version >= new Version(6, 2, 8400, 0))
                NativeMethods.DwmpActivateLivePreview(peekOn, Handle, _topmostHandle, peekType, UIntPtr.Zero);
            else
                NativeMethods.DwmpActivateLivePreview(peekOn, Handle, _topmostHandle, peekType);
        }

        NativeMethods.WINDOWPLACEMENT GetPlacement()
        {
            NativeMethods.WINDOWPLACEMENT placement = new NativeMethods.WINDOWPLACEMENT();
            placement.Length = Marshal.SizeOf(placement);
            NativeMethods.GetWindowPlacement(this.Handle, ref placement);
            return placement;
        }

        public void Show()
        {
            //Debug.WriteLine("Show()");
            //var placement = GetPlacement();

            /*if (IsMinimized)
                NativeMethods.ShowWindow(Handle, NativeMethods.SwShow);*/

            NativeMethods.SetForegroundWindow(Handle);


            if (IsMinimized)
            {
                if (IsMaximized)
                    Maximize();
                else
                    Restore();
            }

            NativeMethods.ShowWindowAsync(Handle, NativeMethods.SwShow);

            /*else if ((placement.ShowCmd == NativeMethods.SwMinimize) || (placement.ShowCmd == NativeMethods.SwForceMinimize) || (placement.ShowCmd == NativeMethods.SwShowMinimized) || (placement.ShowCmd == NativeMethods.SwShowMinimizedNoActive))
                NativeMethods.ShowWindow(Handle, NativeMethods.SwShow);*/
        }

        public void ShowSystemMenu(IntPtr callerWindowHandle)
        {
            Debug.WriteLine("ShowSystemMenu");
            IntPtr menu = NativeMethods.GetSystemMenu(Handle, false);
            uint command = NativeMethods.TrackPopupMenuEx(menu, 0x0000/* | 0x0100*/, 100, 100, Handle, IntPtr.Zero);
            if (command != 0)
                Debug.WriteLine("PostMessage output: " + NativeMethods.PostMessage(Handle, 0x0112, new IntPtr(command), new IntPtr(0xF090)).ToString());
        }
    }

    /*internal class ProcessWindowMonitorForm : Form
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool RegisterShellHookWindow(IntPtr hWnd);

        int HShellWindowCreated = 1;
        int HShellWindowDestroyed = 2;
        int HShellActivateShellWindow = 3;
        int HShellWindowActivated = 4;
        int HShellGetMinRect = 5;
        int HShellRedraw = 6;
        int HShellTaskMan = 7;
        int HShellLanguage = 8;
        int HShellAccessibilityState = 11;
        int HShellAppCommand = 12;

        private readonly uint notificationMsg;
        public delegate void EventHandler(object sender, string data);
        public event EventHandler WindowEvent;
        protected virtual void OnWindowEvent(string data)
        {
            WindowEvent?.Invoke(this, data);
        }

        internal ProcessWindowMonitorForm()
        {
            notificationMsg = RegisterWindowMessage("SHELLHOOK");
            RegisterShellHookWindow(this.Handle);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == notificationMsg)
            base.WndProc(ref m);

            if (m.WParam.ToInt32() == HShellWindowDestroyed)
            {
                ProcessWindow.RaiseWindowClosed(m.HWnd);
            }
            else if (m.WParam.ToInt32() == HShellWindowActivated)
            {
                ProcessWindow.ActiveWindow = new ProcessWindow(m.HWnd);
            }
            else if (m.WParam.ToInt32() == HShellWindowCreated)
            {
                //ProcessWindow.RaiseWindowOpened(m.HWnd);
            }
        }
    }*/

    public class WindowEventArgs : EventArgs
    {
        public WindowEventArgs(ProcessWindow window)
        {
            Window = window;
        }

        public ProcessWindow Window { get; }
        public bool IsVisibilityChange { get; set; } = false;
    }
}
