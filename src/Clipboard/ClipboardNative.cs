
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace Clipboard
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

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GlobalLock(IntPtr hMem);


        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        public static extern IntPtr GlobalSize(IntPtr handle);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool CloseClipboard();


        [DllImport("kernel32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        public static extern int GlobalSize(HandleRef handle);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GlobalUnlock(IntPtr hMem);

    }

    [Flags]
    public enum FileDescriptorFlags : uint
    {
        FD_CLSID = 0x00000001,
        FD_SIZEPOINT = 0x00000002,
        FD_ATTRIBUTES = 0x00000004,
        FD_CREATETIME = 0x00000008,
        FD_ACCESSTIME = 0x00000010,
        FD_WRITESTIME = 0x00000020,
        FD_FILESIZE = 0x00000040,
        FD_PROGRESSUI = 0x00004000,
        FD_LINKUI = 0x00008000,
        FD_UNICODE = 0x80000000 //Windows Vista and later
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct FILEDESCRIPTOR
    {
        public FileDescriptorFlags dwFlags;
        public Guid clsid;
        public System.Drawing.Size sizel;
        public System.Drawing.Point pointl;
        public uint dwFileAttributes;
        public FILETIME ftCreationTime;
        public FILETIME ftLastAccessTime;
        public FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;

        public override string ToString()
        {
            UInt64 size = nFileSizeHigh << 32 | nFileSizeLow;
            return $"{cFileName} ({size} bytes)";
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct FILEGROUPDESCRIPTOR
    {
        public uint cItems;
        public IntPtr fgd;
    }

    public struct DROPFILES
    {
        public uint pFiles;
        public POINT pt;
        [MarshalAs(UnmanagedType.Bool)]
        public bool fNC;
        [MarshalAs(UnmanagedType.Bool)]
        public bool fWide;
    }
}



