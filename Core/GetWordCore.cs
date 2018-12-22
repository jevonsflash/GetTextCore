using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;



namespace Core
{
    public class GetWordCore
    {
        public event EventHandler<string> OnTextChanged;


        public delegate int NotifyCallBack(int wParam, int lParam);

        private const string LICENSEID = "{00000000-0000-0000-0000-000000000000}";

        private const int MOD_ALT = 1;

        private const int MOD_CONTROL = 2;

        private const int MOD_SHIFT = 4;

        private const int MOD_WIN = 8;

        private const int VK_LBUTTON = 1;

        private const int VK_RBUTTON = 2;

        private const int VK_MBUTTON = 4;

        private const int WM_LBUTTONDBLCLK = 515;

        private const int SW_HIDE = 0;

        private const int SW_SHOW = 5;

        private bool bGetWordUnloaded;

        private NotifyCallBack callbackHighlightReady;

        private NotifyCallBack callbackMouseMonitor;

        private IContainer components;


        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out Point lpPoint);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int ShowWindow(IntPtr hwnd, int nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr WindowFromPoint(Point point);

        [DllImport("kernel32.dll")]
        public static extern void Sleep(int uMilliSec);

        [DllImport("kernel32.dll")]
        public static extern int FreeLibrary(int hLibModule);

        [DllImport("kernel32.dll")]
        public static extern int GetModuleHandle(string lpModuleName);

        [DllImport("GetWord.dll")]
        public static extern void SetLicenseID([MarshalAs(UnmanagedType.BStr)] string szLicense);

        [DllImport("GetWord.dll")]
        public static extern void SetNotifyWnd(int hWndNotify);

        [DllImport("GetWord.dll")]
        public static extern void UnSetNotifyWnd(int hWndNotify);

        [DllImport("GetWord.dll")]
        public static extern void SetDelay(int uMilliSec);

        [DllImport("GetWord.dll")]
        public static extern bool EnableCursorCapture(bool bEnable);

        [DllImport("GetWord.dll")]
        public static extern bool EnableHotkeyCapture(bool bEnable, int fsModifiers, int vKey);

        [DllImport("GetWord.dll")]
        public static extern bool EnableHighlightCapture(bool bEnable);

        [DllImport("GetWord.dll")]
        public static extern bool GetString(int x, int y, [MarshalAs(UnmanagedType.BStr)] out string str, out int nCursorPos);

        [DllImport("GetWord.dll")]
        public static extern bool GetString2(int x, int y, [MarshalAs(UnmanagedType.BStr)] out string str, out int nCursorPos, out int left, out int top, out int right, out int bottom);

        [DllImport("GetWord.dll")]
        public static extern bool FreeString([MarshalAs(UnmanagedType.BStr)] out string str);

        [DllImport("GetWord.dll")]
        public static extern bool GetRectString(int hWnd, int left, int top, int right, int bottom, [MarshalAs(UnmanagedType.BStr)] out string str);

        [DllImport("GetWord.dll")]
        public static extern int GetRectStringPairs(int hWnd, int left, int top, int right, int bottom, [MarshalAs(UnmanagedType.BStr)] out string str, [MarshalAs(UnmanagedType.BStr)] out string rectList);

        [DllImport("GetWord.dll")]
        public static extern int GetPointStringPairs(int x, int y, [MarshalAs(UnmanagedType.BStr)] out string str, [MarshalAs(UnmanagedType.BStr)] out string rectList);

        [DllImport("GetWord.dll")]
        public static extern bool FreePairs([MarshalAs(UnmanagedType.BStr)] out string str, [MarshalAs(UnmanagedType.BStr)] out string rectList);

        [DllImport("GetWord.dll")]
        public static extern bool GetPairItem(int totalCount, [MarshalAs(UnmanagedType.BStr)] out string str, [MarshalAs(UnmanagedType.BStr)] out string rectList, int index, [MarshalAs(UnmanagedType.BStr)] out string substr, out int substrLen, out int substrLeft, out int substrTop, out int substrRight, out int substrBottom);

        [DllImport("GetWord.dll")]
        public static extern bool GetHighlightText(int hWnd, [MarshalAs(UnmanagedType.BStr)] out string str);

        [DllImport("GetWord.dll")]
        public static extern bool GetHighlightText2(int x, int y, [MarshalAs(UnmanagedType.BStr)] out string str);

        [DllImport("GetWord.dll")]
        public static extern bool SetCaptureReadyCallback(NotifyCallBack callback);

        [DllImport("GetWord.dll")]
        public static extern bool RemoveCaptureReadyCallback(NotifyCallBack callback);

        [DllImport("GetWord.dll")]
        public static extern bool SetHighlightReadyCallback(NotifyCallBack callback);

        [DllImport("GetWord.dll")]
        public static extern bool RemoveHighlightReadyCallback(NotifyCallBack callback);

        [DllImport("GetWord.dll")]
        public static extern bool SetMouseMonitorCallback(NotifyCallBack callback);

        [DllImport("GetWord.dll")]
        public static extern bool RemoveMouseMonitorCallback(NotifyCallBack callback);


        private string _textHighlight;

        public string TextHighlight
        {
            get { return _textHighlight; }
            set { _textHighlight = value; }
        }


        private bool _checkHighlight;

        public bool CheckHighlight
        {
            get { return _checkHighlight; }
            set { _checkHighlight = value; }
        }


        public void Load()
        {
            SetLicenseID("{00000000-0000-0000-0000-000000000000}");
            SetDelay(300);
            callbackHighlightReady = OnHighlightReady;
            SetHighlightReadyCallback(callbackHighlightReady);
            callbackMouseMonitor = OnMouseMonitor;
            SetMouseMonitorCallback(callbackMouseMonitor);
        }

        public void Unload()
        {
            Debug.WriteLine("Form1_FormClosing");
            bGetWordUnloaded = true;
            RemoveHighlightReadyCallback(callbackHighlightReady);
            RemoveMouseMonitorCallback(callbackMouseMonitor);
        }

        public void Enable(bool isEnable)
        {
            bool flag = false;
            flag = (isEnable ? true : false);
            EnableHighlightCapture(flag);
        }

        private int OnHighlightReady(int wParam, int lParam)
        {
            if (bGetWordUnloaded)
            {
                return 1;
            }
            int result = 0;
            if (bGetWordUnloaded)
            {
                return result;
            }
            int x = wParam & 0xFFFF;
            int y = (int)((wParam & 4294901760u) >> 16);
            string str = "";
            if (GetHighlightText2(x, y, out str))
            {
                TextHighlight = str;
                this.OnTextChanged?.Invoke(this, TextHighlight);
                FreeString(out str);
                return 0;
            }
            TextHighlight = "";

            return 1;
        }

        private int OnMouseMonitor(int wParam, int lParam)
        {
            if (bGetWordUnloaded)
            {
                return 1;
            }
            int result = 0;
            if (bGetWordUnloaded)
            {
                return result;
            }
            if (wParam == 515)
            {
                result = OnHighlightReady(wParam, lParam);
            }
            return result;
        }

        protected void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }
            this.Dispose(disposing);
        }
    }
}
