// © Copyright 2020 Levit & James, Inc.

using System;
using System.Runtime.InteropServices;

namespace LevitJames.Win32
{
    [StructLayout(LayoutKind.Sequential)]
    public struct WindowMessage
    {

        public WindowMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam)
        {
            HWnd = hWnd;
            Msg = msg;
            WParam = wParam;
            LParam = lParam;
            Result = IntPtr.Zero;
        }

        public IntPtr HWnd { get; }

        public int Msg { get; }

        public IntPtr WParam { get; set; }

        public IntPtr LParam { get; set; }

        public IntPtr Result { get; set; }

        public object GetLParam(Type cls) => Marshal.PtrToStructure(LParam, cls);

        public override bool Equals(object other)
        {
            if (!(other is WindowMessage))
                return false;

            var message = (WindowMessage) other;
            return HWnd == message.HWnd && Msg == message.Msg && WParam == message.WParam && LParam == message.LParam && Result == message.Result;
        }

        public static bool operator !=(WindowMessage a, WindowMessage b) => !a.Equals(b);

        public static bool operator ==(WindowMessage a, WindowMessage b) => a.Equals(b);

        public override int GetHashCode() => Msg | ((int) HWnd << 4);

        public override string ToString()
        {

#if DEBUG_WIN_MSG
            return WindowMessages.ToString(this));
#else
            return base.ToString();
#endif

        }
    }
}