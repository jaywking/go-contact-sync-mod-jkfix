using System;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    internal class NativeMethods
    {
        #region API Constants
        public const int HWND_BROADCAST = 0xffff;
        public static readonly int WM_GCSM_SHOWME = RegisterWindowMessage("WM_GCSM_SHOWME");
        //public const int VER_NT_WORKSTATION = 0x0000001;

        // Fix for WinXP and older systems, that do not continue with shutdown until all programs have closed
        // FormClosing would hold system shutdown, when it sets the cancel to true
        public const int WM_QUERYENDSESSION = 0x11;

        #endregion

        #region Extern Functions Declaration

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int RegisterWindowMessage(string message);

        [DllImport("ole32.dll")]
        public static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        #endregion
    }
}
