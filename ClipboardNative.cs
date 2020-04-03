
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace clipview
{
    internal static class ClipboardNative
    {
        [DllImport("user32", SetLastError = true)]
        public static extern int CountClipboardFormats();

        [DllImport("user32", SetLastError = true)]
        public static extern uint EnumClipboardFormats(uint format);

        [DllImport("user32", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsClipboardFormatAvailable(uint format);

        [DllImport("user32", SetLastError = true, CharSet = CharSet.Unicode)]
        // private static extern unsafe int GetClipboardFormatName(uint format, [MarshalAs(UnmanagedType.LPWStr)] [Out] char* lpszFormatName, int cchMaxCount);
        private static extern unsafe int GetClipboardFormatName(uint format, [MarshalAs(UnmanagedType.LPWStr)] [Out] StringBuilder lpszFormatName, int cchMaxCount);
        public static string GetClipboardFormatName(uint format)
        {
            if (format < (uint)StandardClipboardFormats.StartOfPrivateRange)
            {
                return ((StandardClipboardFormats)format).ToString();
            }
            StringBuilder sb = new StringBuilder(100);
            int ret = GetClipboardFormatName(format, sb, sb.Capacity);
            if (ret < sb.Capacity)
            {
                return sb.ToString();
            }
            else
            {
                throw new Exception("StringBuilder not big enough");
            }
        }



        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        public static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GlobalLock(IntPtr hMem);


        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalSize(IntPtr handle);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool CloseClipboard();


        [DllImport("kernel32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        public static extern int GlobalSize(HandleRef handle);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GlobalUnlock(IntPtr hMem);

    }
}



