using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Microsoft.WindowsAPICodePack.Shell.Interop.Common
{
    internal static class WndProcRetHookNativeMethods
    {
        /// <summary>
        /// Specific declaration to use with WH_CALLWNDPROCRET.
        /// See https://msdn.microsoft.com/en-us/library/windows/desktop/ms644990(v=vs.85).aspx for the doc.
        /// </summary>
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(HookType idHook, CallWndRetProc lpfn, IntPtr hMod, uint dwThreadId);

        public static IntPtr SetWindowsHookEx(CallWndRetProc callback)
            => SetWindowsHookEx(HookType.WH_CALLWNDPROCRET, callback, IntPtr.Zero, WindowNativeMethods.GetCurrentThreadId());

        /// <summary>
        /// Passes the hook information to the next hook procedure in the current hook chain.
        /// A hook procedure can call this function either before or after processing the hook information. 
        /// </summary>
        /// <param name="hhk">
        /// This parameter is ignored.
        /// </param>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, CWPRETSTRUCT lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);
    }

    /// <summary>
    /// Callback for SetWindowsHookEx, when using the WH_CALLWNDPROCRET hook.
    /// </summary>
    internal delegate IntPtr CallWndRetProc(int code, IntPtr wParam, CWPRETSTRUCT lParam);

    [StructLayout(LayoutKind.Sequential)]
    internal class CWPRETSTRUCT
    {
        public IntPtr lResult;
        public IntPtr lParam;
        public IntPtr wParam;
        public IntPtr message;
        public IntPtr hWnd;
    }

    internal enum HookType
    {
        WH_MIN = -1,
        WH_MSGFILTER = -1,
        WH_JOURNALRECORD = 0,
        WH_JOURNALPLAYBACK = 1,
        WH_KEYBOARD = 2,
        WH_GETMESSAGE = 3,
        WH_CALLWNDPROC = 4,
        WH_CBT = 5,
        WH_SYSMSGFILTER = 6,
        WH_MOUSE = 7,
        WH_HARDWARE = 8,
        WH_DEBUG = 9,
        WH_SHELL = 10,
        WH_FOREGROUNDIDLE = 11,
        WH_CALLWNDPROCRET = 12,
        WH_KEYBOARD_LL = 13,
        WH_MOUSE_LL = 14
    }
}
