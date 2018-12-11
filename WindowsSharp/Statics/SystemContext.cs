using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsSharp.Statics
{
    public class SystemContext
    {
        public static SystemContext Instance = new SystemContext();

        private SystemContext()
        {

        }

        /// <summary>
        /// Gets the avatar of the currently logged in user.
        /// </summary>
        /// <value>
        /// A brush that paints the avatar image.
        /// </value>
        public System.Drawing.Image UserAvatar
        {
            get
            {
                var sb = new StringBuilder(1000);
                NativeMethods.GetUserTilePath(null, 0x80000000, sb, sb.Capacity);
                return System.Drawing.Image.FromFile(sb.ToString());
            }
        }

        /// <summary>
        /// Gets the state of the desktop window manager.
        /// </summary>
        /// <value>
        /// A boolean representing whether the desktop window manager is enabled or not.
        /// </value>
        public bool IsDesktopWindowManagerEnabled
        {
            get
            {
                if (Environment.OSVersion.Version.Major >= 6)
                    return NativeMethods.DwmIsCompositionEnabled();
                else
                    return false;
            }
        }

        /// <summary>
        /// Gets the name of the currently applied visual style.
        /// </summary>
        /// <value>
        /// A string that is the current visual style's name.
        /// </value>
        public string CurrentVisualStyleName
        {
            get
            {
                var visualStyleName = new StringBuilder(0x200);
                NativeMethods.GetCurrentThemeName(visualStyleName, visualStyleName.Capacity, null, 0, null, 0);
                return visualStyleName.ToString();
            }
        }

        /// <summary>
        /// Gets the currently applied accent color.
        /// </summary>
        /// <value>
        /// A brush that can paint the windows accent color.
        /// </value>
        public System.Drawing.Color WindowGlassColor
        {
            get
            {
                if (IsDesktopWindowManagerEnabled)
                {
                    ///https://stackoverflow.com/questions/13660976/get-the-active-color-of-windows-8-automatic-color-theme
                    NativeMethods.DwmColorizationParams parameters = new NativeMethods.DwmColorizationParams();
                    NativeMethods.DwmGetColorizationParameters(ref parameters);
                    var coloures = parameters.ColorizationColor.ToString("X");
                    while (coloures.Length < 8)
                    {
                        coloures = "0" + coloures;
                    }
                    var alphaBase = Int32.Parse(coloures.Substring(0, 2), NumberStyles.HexNumber);
                    var alphaMultiplier = ((Double)(parameters.ColorizationColorBalance + parameters.ColorizationBlurBalance)) / 128;
                    var alpha = (Byte)(alphaBase * alphaMultiplier);
                    System.Diagnostics.Debug.WriteLine("balance over 255: " + (((Double)(parameters.ColorizationColorBalance)) / 255) + "\nalpha: " + alpha);
                    return System.Drawing.Color.FromArgb(alpha, byte.Parse(coloures.Substring(2, 2), NumberStyles.HexNumber), byte.Parse(coloures.Substring(4, 2), NumberStyles.HexNumber), byte.Parse(coloures.Substring(6, 2), NumberStyles.HexNumber));
                }
                else if (Environment.OSVersion.Version.Major <= 5)
                    return System.Drawing.Color.FromArgb(0xFF, 0, 0x53, 0xE1);
                else if (Environment.OSVersion.Version.Major == 6)
                {
                    if (Environment.OSVersion.Version.Minor == 4)
                        return System.Drawing.Color.FromArgb(0xFF, 0x65, 0xC0, 0xF2);
                    if (Environment.OSVersion.Version.Minor == 3)
                        return System.Drawing.Color.FromArgb(0xFF, 0xF0, 0xC8, 0x69);
                    else if (Environment.OSVersion.Version.Minor == 2)
                        return System.Drawing.Color.FromArgb(0xFF, 0x6B, 0xAD, 0xF6);
                    else
                        return System.Drawing.Color.FromArgb(0xFF, 0xB9, 0xD1, 0xEA);
                }
                else
                {
                    return System.Drawing.Color.FromArgb(0xFF, 0x18, 0x83, 0xD7);
                }
            }
        }

        /// <summary>
        /// Locks the computer.
        /// </summary>
        public void LockUserAccount()
        {
            NativeMethods.LockWorkStation();
        }

        /// <summary>
        /// Signs the user out of the computer.
        /// </summary>
        public void SignOut()
        {
            NativeMethods.ExitWindowsEx(NativeMethods.ExitWindowsAction.Force, 0);
        }

        /// <summary>
        /// Puts the computer to sleep.
        /// </summary>
        public void SleepSystem()
        {
            NativeMethods.SetSuspendState(false, true, true);
        }

        /// <summary>
        /// Shuts down the computer.
        /// </summary>
        public void ShutDownSystem()
        {
            NativeMethods.ExitWindowsEx(NativeMethods.ExitWindowsAction.Shutdown, 0);
        }

        /// <summary>
        /// Restarts the computer.
        /// </summary>
        public void RestartSystem()
        {
            NativeMethods.ExitWindowsEx(NativeMethods.ExitWindowsAction.Reboot, 0);
        }
    }
}
