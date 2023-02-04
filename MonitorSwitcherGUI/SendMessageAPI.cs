using System;
using System.Runtime.InteropServices;

namespace MonitorSwitcherGUI
{
    public class SendMessageAPI
    {
        public const int HWND_BROADCAST = 0xFFFF;
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_MONITORPOWER = 0xf170;

        public const int MONITOR_ON = -1;
        public const int MONITOR_OFF = 2;
        public const int MONITOR_STANBY = 1;

        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32")]
        public static extern IntPtr GetConsoleWindow();
    }
}
