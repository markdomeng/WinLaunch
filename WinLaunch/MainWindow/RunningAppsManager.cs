using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WinLaunch
{

    public partial class RunningAppsManager
    {
        public const int GCL_HICONSM = -34;
        public const int GCL_HICON = -14;

        public const int ICON_SMALL = 0;
        public const int ICON_BIG = 1;
        public const int ICON_SMALL2 = 2;
        private int SW_SHOWDEFAULT = 10;
        private int SW_SHOW = 5;
        private int SW_RESTORE = 9;
        private int GA_ROOTOWNER = 3;

        const int GWL_EXSTYLE = -20;
        const uint DWMWA_CLOAKED = 14;
        const uint DWM_CLOAKED_SHELL = 0x00000002;
        const uint WS_EX_TOOLWINDOW = 0x00000080;
        const uint WS_EX_TOPMOST = 0x00000008;
        const uint WS_EX_NOACTIVATE = 0x08000000;

        public const int WM_GETICON = 0x7F;
        public bool IsLoaded { get; private set; }
        public bool IsInitialized { get; private set; }

        BackgroundWorker backgroundWorker;

        // This is the Interop/WinAPI that will be used
        [DllImport("user32.dll", EntryPoint = "GetAncestor")]
        public static extern IntPtr GetAncestor(IntPtr hWnd, int flags);

        [DllImport("user32.dll", EntryPoint = "GetLastActivePopup")]
        public static extern IntPtr GetLastActivePopup(IntPtr hWnd);
        
        [DllImport("dwmapi.dll", EntryPoint = "DwmGetWindowAttribute")]
        public static extern uint DwmGetWindowAttribute(IntPtr hWnd, uint dwAttribute, out uint pvAttribute, int cbAttribute);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongA")]
        public static extern long GetWindowLongA(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fUnknown);

        [DllImport("user32.dll", EntryPoint = "GetClassLong")]
        public static extern uint GetClassLongPtr32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetClassLongPtr")]
        public static extern IntPtr GetClassLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "GetWindowText",
        ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd,
            StringBuilder lpWindowText, int nMaxCount);

        [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows",
        ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool EnumDesktopWindows(IntPtr hDesktop,
            EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

        // Define the callback delegate's type.
        private delegate bool EnumDelegate(IntPtr hWnd, int lParam);

        public List<RunningWindow> RunningWindows { get; set; }

        public RunningAppsManager()
        {
            RunningWindows = new List<RunningWindow>();

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;

            IsLoaded = false;
            IsInitialized = false;
            //Initialize();
        }

        public void Initialize()
        {
            if (!backgroundWorker.IsBusy)
            {
                backgroundWorker.RunWorkerAsync();
            }
        }
        public void Load()
        {
            //if(!backgroundWorker.IsBusy)
            //{
            //    backgroundWorker.RunWorkerAsync();
            //}
            RunningWindows.Clear();
            EnumDesktopWindows(IntPtr.Zero, FilterCallback, IntPtr.Zero);
            IsLoaded = true;

        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            IsLoaded = true;
            IsInitialized = true;
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            IsLoaded = false;
            RunningWindows.Clear();
            EnumDesktopWindows(IntPtr.Zero, FilterCallback, IntPtr.Zero);
        }

        public void BringToFront(IntPtr sender)
        {
            SwitchToThisWindow(sender, true);
        }

        // We use this function to filter windows.
        // This version selects visible windows that have titles.
        public struct RunningWindow
        {
            public string Title { get; set; }
            public BitmapSource Icon { get; set; }
            public IntPtr Handle { get; set; }
        }
        private bool FilterCallback(IntPtr hWnd, int lParam)
        {
            // Get the window's title.
            StringBuilder sb_title = new StringBuilder(1024);
            int length = GetWindowText(hWnd, sb_title, sb_title.Capacity);
            string title = sb_title.ToString();
            // If the window is visible and has a title, save it.
            if (IsAltTabWindow(hWnd) &&
                string.IsNullOrEmpty(title) == false)
            {
                RunningWindow window = new RunningWindow
                {
                    Title = title,
                    Handle = hWnd,
                    Icon = GetAppIcon(hWnd)
                };
                RunningWindows.Add(window);
            }

            // Return true to indicate that we
            // should continue enumerating windows.
            return true;
        }

        private bool IsAltTabWindow(IntPtr hWnd)
        {
            // Start at the root owner
            // The window must be visible
            if (!IsWindowVisible(hWnd))
                return false;

            // The window must be a root owner
            if (GetAncestor(hWnd, GA_ROOTOWNER) != hWnd)
                return false;

            // The window must not be cloaked by the shell
            DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out uint cloaked, sizeof(uint));
            if (cloaked == DWM_CLOAKED_SHELL)
                return false;

            // The window must not have the extended style WS_EX_TOOLWINDOW
            long style = GetWindowLongA(hWnd, GWL_EXSTYLE);
            if ((style & WS_EX_TOOLWINDOW) != 0)
                return false;

            return true;
        }

        public BitmapSource GetAppIcon(IntPtr hwnd)
        {
            IntPtr iconHandle = SendMessage(hwnd, WM_GETICON, ICON_SMALL2, 0);
            if (iconHandle == IntPtr.Zero)
                iconHandle = SendMessage(hwnd, WM_GETICON, ICON_SMALL, 0);
            if (iconHandle == IntPtr.Zero)
                iconHandle = SendMessage(hwnd, WM_GETICON, ICON_BIG, 0);
            if (iconHandle == IntPtr.Zero)
                iconHandle = GetClassLongPtr(hwnd, GCL_HICON);
            if (iconHandle == IntPtr.Zero)
                iconHandle = GetClassLongPtr(hwnd, GCL_HICONSM);

            if (iconHandle == IntPtr.Zero)
                return null;

            Icon icn = Icon.FromHandle(iconHandle);

            return Convert(icn.ToBitmap());
        }

        public IntPtr GetClassLongPtr(IntPtr hWnd, int nIndex)
        {
            if (IntPtr.Size > 4)
                return GetClassLongPtr64(hWnd, nIndex);
            else
                return new IntPtr(GetClassLongPtr32(hWnd, nIndex));
        }
        public BitmapSource Convert(System.Drawing.Bitmap bitmap)
        {
            var bitmapData = bitmap.LockBits(
                new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height),
                System.Drawing.Imaging.ImageLockMode.ReadOnly, bitmap.PixelFormat);

            var bitmapSource = BitmapSource.Create(
                bitmapData.Width, bitmapData.Height,
                bitmap.HorizontalResolution, bitmap.VerticalResolution,
                PixelFormats.Pbgra32, null,
                bitmapData.Scan0, bitmapData.Stride * bitmapData.Height, bitmapData.Stride);

            bitmap.UnlockBits(bitmapData);

            return bitmapSource;
        }

    }

}
