using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace Clipboard
{
    public class DataObject
    {
        private readonly Dictionary<string, StandardClipboardFormats> formats = new Dictionary<string, StandardClipboardFormats>();

        private readonly IDataObject inner;

        public bool UseString { get; set; }
        public bool UseAscii { get; set; }
        public DataObject()
        {
            formats.Clear();
            for (int i = 0; i < 3 && inner == null; i++)
            {
                try
                {
                    inner = OleClipboardNative.OleGetClipboard();
                }
                catch (COMException c)
                {
                    if ((uint)c.HResult == 0x800401D0u) // CLIPBRD_E_CANT_OPEN
                    {
                        System.Threading.Thread.Sleep(10);
                    }
                }
            }
            if (inner == null) { throw new COMException("timeout"); }

            var enumFmt = inner.EnumFormatEtc(DATADIR.DATADIR_GET);
            FORMATETC[] fmtEtc = new FORMATETC[1];
            int[] fetched = new int[1];
            while (enumFmt.Next(1, fmtEtc, fetched) == 0)
            {
                var fmt = fmtEtc[0].cfFormat;
                string name = ClipboardNative.GetClipboardFormatName((uint)fmt);
                RegisterFormat(name, (uint)fmt);
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint RegisterClipboardFormat(string formatName);
        private void RegisterFormat(string name, uint fmt)
        {
            if (fmt < (uint)StandardClipboardFormats.StartOfPrivateRange)
            {
                formats[name] = (StandardClipboardFormats)fmt;
            }
            else if (!formats.ContainsKey(name))
            {
                uint value = RegisterClipboardFormat(name);
                formats[name] = (StandardClipboardFormats)value;
            }
        }

        public IEnumerable<string> GetClipboardFormats()
        {
            return formats.Keys;
        }

        /// <summary>
        /// Gets the data on the clipboard in the format specified
        /// </summary>
        private byte[] InternalGetData(string format)
        {
            if (!GetClipboardFormats().Contains(format))
            {
                throw new KeyNotFoundException();
            }
            uint fmt = (uint)formats[format];


            //Get pointer to clipboard data in the selected format
            IntPtr ClipboardDataPointer = ClipboardNative.GetClipboardData(fmt);

            //Do a bunch of crap necessary to copy the data from the memory
            //the above pointer points at to a place we can access it.
            IntPtr Length = ClipboardNative.GlobalSize(ClipboardDataPointer);
            IntPtr gLock = ClipboardNative.GlobalLock(ClipboardDataPointer);

            //Init a buffer which will contain the clipboard data
            byte[] Buffer = new byte[(int)Length];

            //Copy clipboard data to buffer
            Marshal.Copy(gLock, Buffer, 0, (int)Length);
            return Buffer;
        }

        public object Win32GetData(string format)
        {
            return new MemoryStream(InternalGetData(format));
        }
        public object GetData(string format, int index = -1)
        {
            if (format == "FileContents" && index == -1)
            {
                var files = (FILEDESCRIPTOR[])GetData("FileGroupDescriptorW");
                var streams = new FileContentsStream[files.Length];
                for (int i = 0; i < files.Length; i++)
                {
                    streams[i] = new FileContentsStream((IStream)GetData(format, i));
                    string name = streams[i].FileName;
                    if (files[i].cFileName != name)
                    {
                        throw new Exception("Something went wrong");
                    }
                }
                return streams;
            }
            using (var medium = OleGetData(format, index))
            {
                medium.UseAscii = UseAscii;
                switch (format)
                {
                    case "Text":
                    case "UnicodeText":
                    case "OemText":
                    case "Locale":
                    case "FileName":
                    case "FileNameW":
                    case "HTML Format":
                    case "UniformResourceLocatorW":
                    case "Csv":
                        return medium.GetString();
                    case "FileGroupDescriptorW":
                        return medium.GetFileGroupDescriptor();
                    case "FileContents":
                        return medium.GetStream();
                    case "PNG":
                    default:
                        if (UseString) { return medium.GetString(); }
                        else { return new MemoryStream(medium.GetByteArray()); }
                }
            }
        }
        private TYMED GetPreferredMediumForFormat(string format)
        {
            switch (format)
            {
                case "Text":
                case "UnicodeText":
                case "OemText":
                case "Locale":
                    return TYMED.TYMED_HGLOBAL;
                case "FileName":
                case "FileNameW":
                case "FileGroupDescriptorW":
                case "UniformResourceLocatorW":
                    return TYMED.TYMED_HGLOBAL;
                case "FileContents":
                case "HTML Format":
                case "Csv":
                    return TYMED.TYMED_ISTREAM;
                case "Bitmap":
                    return TYMED.TYMED_GDI;
                case "DeviceIndependentBitmap":
                case "DeviceIndependentBitmapW":
                case "DeviceIndependentBitmapV5":
                case "PNG":
                    return TYMED.TYMED_HGLOBAL;
            }
            return TYMED.TYMED_HGLOBAL;
        }

        public class NativeStgMedium : IDisposable
        {
            STGMEDIUM medium;
            public string Format { get; private set; }
            public bool UseAscii { get; set; }

            public NativeStgMedium(STGMEDIUM medium, string format) { this.medium = medium; Format = format; }

            [DllImport("ole32.dll")]
            static extern void ReleaseStgMedium(ref STGMEDIUM medium);
            public unsafe void Dispose()
            {
                ReleaseStgMedium(ref medium);
            }

            public IStream GetStream()
            {
                if (medium.tymed == TYMED.TYMED_ISTREAM)
                {
                    var stream = (IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
                    return stream;
                }
                throw new InvalidComObjectException();
            }


            public byte[] GetByteArray(int size = -1)
            {
                if (medium.tymed == TYMED.TYMED_HGLOBAL)
                {
                    IntPtr Length = ClipboardNative.GlobalSize(medium.unionmember);
                    if (size != -1)
                    {
                        if (size != (int)Length)
                        {
                            throw new ArgumentException("Wrong size");
                        }
                    }
                    IntPtr gLock = ClipboardNative.GlobalLock(medium.unionmember);
                    try
                    {
                        //Init a buffer which will contain the clipboard data
                        byte[] Buffer = new byte[(int)Length];

                        //Copy clipboard data to buffer
                        Marshal.Copy(gLock, Buffer, 0, (int)Length);
                        return Buffer;
                    }
                    finally { ClipboardNative.GlobalUnlock(gLock); }
                }
                else if (medium.tymed == TYMED.TYMED_GDI)
                {
                    using (var bitmap = new NativeBitmap(medium.unionmember))
                    {
                        return bitmap.Bitmap.Clone() as byte[];
                    }
                }
                else
                {
                    throw new FormatException();
                }
            }

            public uint GetDWORD()
            {
                if (medium.tymed == TYMED.TYMED_HGLOBAL)
                {
                    byte[] bytes = GetByteArray();
                    return BitConverter.ToUInt32(bytes, 0);
                }
                else
                {
                    throw new InvalidComObjectException();
                }
            }

            public FILEDESCRIPTOR[] GetFileGroupDescriptor()
            {
                if (medium.tymed == TYMED.TYMED_HGLOBAL)
                {
                    IntPtr Length = ClipboardNative.GlobalSize(medium.unionmember);
                    IntPtr gLock = ClipboardNative.GlobalLock(medium.unionmember);
                    try
                    {
                        FILEGROUPDESCRIPTOR fgd = Marshal.PtrToStructure<FILEGROUPDESCRIPTOR>(gLock);
                        return fgd.fgd;
                    }
                    finally { ClipboardNative.GlobalUnlock(gLock); }
                }
                else
                {
                    throw new COMException();
                }
            }

            public string GetString()
            {
                if (medium.tymed == TYMED.TYMED_HGLOBAL)
                {
                    byte[] bytes = GetByteArray();
                    switch (Format)
                    {
                        case "Text":
                        case "FileName":
                        case "OemText":
                            return Encoding.ASCII.GetString(bytes, 0, bytes.Length - 1);
                        case "UnicodeText":
                        case "FileNameW":
                            return Encoding.Unicode.GetString(bytes, 0, bytes.Length - 2);
                        case "Locale":
                            return GetDWORD().ToString();
                    }
                    return UseAscii ? Encoding.ASCII.GetString(GetByteArray()) : Encoding.Unicode.GetString(GetByteArray());
                }
                else if (medium.tymed == TYMED.TYMED_FILE)
                {
                    return Marshal.PtrToStringBSTR(medium.unionmember);
                }
                else if (medium.tymed == TYMED.TYMED_ISTREAM)
                {
                    using (var ms = new MemoryStream())
                    {
                        using (var fileStream = new FileContentsStream(GetStream()))
                        {
                            fileStream.SaveToStream(ms);
                            ms.Position = 0;
                            using (var sr = new StreamReader(ms))
                            {
                                return sr.ReadToEnd();
                            }
                        }
                    }
                }
                throw new InvalidComObjectException();
            }
        }

        private static STATSTG Stat(IStream stream)
        {
            stream.Stat(out STATSTG statstg, 0);
            return statstg;
        }
        internal NativeStgMedium OleGetData(string format, int index = -1)
        {
            FORMATETC fmtEtc = new FORMATETC();
            fmtEtc.cfFormat = (short)formats[format];
            fmtEtc.tymed = GetPreferredMediumForFormat(format);
            fmtEtc.lindex = index;
            inner.GetData(ref fmtEtc, out STGMEDIUM medium);
            return new NativeStgMedium(medium, format);
        }
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
        public UInt32 dwFileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
        public UInt32 nFileSizeHigh;
        public UInt32 nFileSizeLow;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct FILEGROUPDESCRIPTOR
    {
        public uint cItems;
        [MarshalAs(UnmanagedType.ByValArray)]
        public FILEDESCRIPTOR[] fgd;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BITMAP
    {
        public int bmType;
        public int bmWidth;
        public int bmHeight;
        public int bmWidthBytes;
        public ushort bmPlanes;
        public ushort bmBitsPixel;
        public IntPtr bmBits;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct BITMAPINFO
    {
        public BITMAPINFOHEADER bmiHeader;
        public byte bmiColors_rgbBlue;
        public byte bmiColors_rgbGreen;
        public byte bmiColors_rgbRed;
        public byte bmiColors_rgbReserved;
    }

    public enum DIB_Color_Mode : uint
    {
        DIB_RGB_COLORS = 0,
        DIB_PAL_COLORS = 1
    }

    public static class Gdi32
    {
        [DllImport("gdi32.dll", CharSet = CharSet.Auto, EntryPoint = "GetObject")]
        public static extern int GetObjectBitmap(IntPtr hObject, int nCount, ref BITMAP lpObject);

        /// <summary>
        ///        Retrieves the bits of the specified compatible bitmap and copies them into a buffer as a DIB using the specified format.
        /// </summary>
        /// <param name="hdc">A handle to the device context.</param>
        /// <param name="hbmp">A handle to the bitmap. This must be a compatible bitmap (DDB).</param>
        /// <param name="uStartScan">The first scan line to retrieve.</param>
        /// <param name="cScanLines">The number of scan lines to retrieve.</param>
        /// <param name="lpvBits">A pointer to a buffer to receive the bitmap data. If this parameter is <see cref="IntPtr.Zero"/>, the function passes the dimensions and format of the bitmap to the <see cref="BITMAPINFO"/> structure pointed to by the <paramref name="lpbi"/> parameter.</param>
        /// <param name="lpbi">A pointer to a <see cref="BITMAPINFO"/> structure that specifies the desired format for the DIB data.</param>
        /// <param name="uUsage">The format of the bmiColors member of the <see cref="BITMAPINFO"/> structure. It must be one of the following values.</param>
        /// <returns>If the lpvBits parameter is non-NULL and the function succeeds, the return value is the number of scan lines copied from the bitmap.
        /// If the lpvBits parameter is NULL and GetDIBits successfully fills the <see cref="BITMAPINFO"/> structure, the return value is nonzero.
        /// If the function fails, the return value is zero.
        /// This function can return the following value: ERROR_INVALID_PARAMETER (87 (0×57))</returns>
        [DllImport("gdi32.dll", EntryPoint = "GetDIBits", SetLastError = true)]
        private static extern int GetDIBits([In] IntPtr hdc, [In] IntPtr hbmp, uint uStartScan, uint cScanLines, [Out] byte[] lpvBits, ref BITMAPINFO lpbi, DIB_Color_Mode uUsage);
        public static byte[] GetDIBits(IntPtr hdc, IntPtr hBitmap, uint bmpHeight, out BITMAPINFO bitmapInfo)
        {
            const int DIB_RGB_COLORS = 0;
            byte[] retVal = null;
            bitmapInfo = new BITMAPINFO();
            bitmapInfo.bmiHeader.biSize = Marshal.SizeOf(bitmapInfo.bmiHeader);
            int success = GetDIBits(hdc, hBitmap, 0, bmpHeight, null, ref bitmapInfo, DIB_RGB_COLORS); // BUGBUG : Won't work for palettize Bitmap
            if (success == 0)
            {
                throw new ExternalException($"Call to Native API 'GetDIBits' failed, error 0x{Marshal.GetLastWin32Error().ToString("x")}");
            }

            retVal = new byte[bitmapInfo.bmiHeader.biSizeImage];
            bitmapInfo.bmiHeader.biCompression = 0; // BI_RGB
            success = GetDIBits(hdc, hBitmap, 0, bmpHeight, retVal, ref bitmapInfo, DIB_RGB_COLORS); // BUGBUG : Won't work for palettize Bitmap
            if (success == 0)
            {
                throw new ExternalException("Second call to Native API 'GetDIBits' failed, error #" + Marshal.GetLastWin32Error().ToString());
            }

            return retVal;
        }

        [DllImport("gdi32.dll")]
        internal static extern bool DeleteObject(IntPtr hObject);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);
    }

}