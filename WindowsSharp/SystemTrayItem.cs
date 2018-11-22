using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsSharp.DiskItems;
using System.Windows.Input;
using System.Diagnostics;
using System.ComponentModel;

namespace WindowsSharp
{
    public class SystemTrayItem : INotifyPropertyChanged
    {
        private void NotifyPropertyChanged(string info)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(info));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        static NativeMethods.NotificationCb _callback = new NativeMethods.NotificationCb();
        static NativeMethods.TrayNotify _notifier = (new NativeMethods.TrayNotify());

        static SystemTrayItem()
        {
            if (Environment.OSVersion.Version >= new Version(6, 2, 8400, 0))
            {
                var handle = default(ulong);


                ((NativeMethods.ITrayNotify)_notifier).RegisterCallback(_callback, out handle);
                //notifier.UnregisterCallback(handle);
            }
            else
            {
                //var notifier = (NativeMethods.ITrayNotifyWin7)(new NativeMethods.TrayNotify());

                ((NativeMethods.ITrayNotifyWin7)_notifier).RegisterCallback(_callback);
                //((NativeMethods.ITrayNotifyWin7)_notifier).RegisterCallback(null);
            }
        }

        public static List<SystemTrayItem> TrayItems
        {
            get
            {
                List<SystemTrayItem> items = new List<SystemTrayItem>();

                var rawItems = _callback.items;
                foreach (NativeMethods.NOTIFYITEM t in rawItems)
                    items.Add(new SystemTrayItem(t));

                return items;
            }
        }

        NativeMethods.NOTIFYITEM _item;

        public enum TrayDisplayMode
        {
            ShowNotifications,
            ShowNever,
            ShowAlways
        }

        TrayDisplayMode _displayMode = TrayDisplayMode.ShowAlways;
        public TrayDisplayMode DisplayMode
        {
            get => _displayMode;
            set
            {
                _displayMode = value;
                NotifyPropertyChanged("DisplayMode");
            }
        }

        public Icon ItemIcon
        {
            get
            {
                Debug.WriteLine("ICON HANDLE: " + _item.hIcon.ToString());
                return Icon.FromHandle(_item.hIcon);
            }
            private set
            {
                NotifyPropertyChanged("ItemIcon");
            }
        }

        public string ToolTipText
        {
            get => _item.pszTip;
            private set
            {
                NotifyPropertyChanged("ToolTipText");
            }
        }

        public DiskItem Executable
        {
            get => new DiskItem(_item.pszExeName);
            private set
            {
                NotifyPropertyChanged("Executable");
            }
        }

        public SystemTrayItem(NativeMethods.NOTIFYITEM item)
        {
            _item = item;

            if (_item.dwPreference == 0)
                DisplayMode = TrayDisplayMode.ShowNotifications;
            else if (_item.dwPreference == 1)
                DisplayMode = TrayDisplayMode.ShowNever;
            else
                DisplayMode = TrayDisplayMode.ShowAlways;

            NotifyPropertyChanged("ItemIcon");
            Debug.WriteLine("Item window handle: " + _item.hWnd.ToString());
        }

        public delegate void CallBack(IntPtr notifyitem);

        public long ActivateTrayItem(uint button)
        {
            if (NativeMethods.IsWindow(_item.hWnd))
            {
                //NativeMethods.PostMessage(_item.hwnd, _item.uCallbackMessage, new IntPtr(_item.id), new IntPtr(NativeMethods.WmRButtonDown));
                //return NativeMethods.SendMessage(_item.hwnd, _item.uCallbackMessage, new IntPtr(_item.id), new IntPtr(NativeMethods.WmRButtonDown)).ToInt32() > 0;
                try
                {
                    //_callback.Notify(NativeMethods.WmRButtonDown, ref _item);
                    Debug.WriteLine("_item.hWnd: " + _item.hWnd.ToString());
                    Debug.WriteLine("_item.hIcon: " + _item.hIcon.ToString());
                    //Debug.WriteLine("_item.uCallbackMessage: " + _item.uCallbackMessage.ToString());
                    /*Debug.WriteLine("_item.reserved: " + _item.reserved.ToString());
                    Debug.WriteLine("_item.reserved2: " + _item.reserved2.ToString());
                    Debug.WriteLine("_item.reserved3: " + _item.reserved3.ToString());
                    Debug.WriteLine("_item.reserved4: " + _item.reserved4.ToString());
                    Debug.WriteLine("_item.reserved5: " + _item.reserved5.ToString());*/
                    /*var win = new Processes.ProcessWindow(_item.hWnd);
                    win.Show();*/
                    //NativeMethods.PostMessage(_item.hWnd, _item.reserved2, new IntPtr(_item.uID), new IntPtr(NativeMethods.WmRButtonDown));
                    /*var data = new NativeMethods.NOTIFYICONDATA()
                    {
                        hWnd = _item.hWnd,
                        uID = Convert.ToInt32(_item.uID),
                        hIcon = _item.hIcon
                    };
                    Debug.WriteLine("data.uCallbackMessage: " + data.uCallbackMessage.ToString());*/
                    IntPtr returnValue = NativeMethods.SendMessage(_item.hWnd, _item.uID, new IntPtr(_item.uID), new UIntPtr(NativeMethods.WmRButtonDown));
                    if (Environment.Is64BitProcess)
                        return returnValue.ToInt64();
                    else
                        return returnValue.ToInt32();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("SystemTrayItem.ActivateTrayItem machine broke:\n" + ex);
                    return 0;
                }
            }
            else
            {
                Debug.WriteLine("Window does not exist: " + _item.hWnd);
                return 0;
            }
        }
    }
}
