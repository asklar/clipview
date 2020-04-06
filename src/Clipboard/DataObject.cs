using System;
using System.Collections.Generic;
using System.IO;
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
            if (inner == null) { throw new TimeoutException(); }

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

        public object GetData(string format, int index = -1)
        {
            try
            {
                return InternalGetData(format, index);
            }
            catch (AccessViolationException)
            {
                return null;
            }
        }

        private object InternalGetData(string format, int index)
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
                    case "Rich Text Format":
                        return medium.GetString();
                    case "FileGroupDescriptorW":
                        return medium.GetFileGroupDescriptor();
                    case "FileContents":
                        return medium.GetStream();
                    case "Drop":
                        return medium.GetDrop();
                    case "PNG":
                        return new MemoryStream(medium.GetByteArray());
                    case "MetafilePicture":
                        return medium.GetMetafile(false);
                    case "EnhancedMetafile":
                        return medium.GetMetafile(true);
                    default:
                        if (UseString) { return medium.GetString(); }
                        else
                        {
                            var bytes = medium.GetByteArray();
                            var ms = new MemoryStream(bytes);
                            switch (ms.Length)
                            {
                                case 8:
                                    return BitConverter.ToInt64(bytes);
                                case 4:
                                    return BitConverter.ToInt32(bytes);
                                case 2:
                                    return BitConverter.ToInt16(bytes);
                                case 1:
                                    return (int)BitConverter.ToChar(bytes);
                                default:
                                    return ms;
                            }
                        }
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
                case "Rich Text Format":
                    return TYMED.TYMED_ISTREAM;
                case "Bitmap":
                    return TYMED.TYMED_GDI;
                case "DeviceIndependentBitmap":
                case "DeviceIndependentBitmapW":
                case "DeviceIndependentBitmapV5":
                case "PNG":
                case "Drop":
                    return TYMED.TYMED_HGLOBAL;
                case "MetafilePicture":
                    return TYMED.TYMED_MFPICT;
                case "EnhancedMetafile":
                    return TYMED.TYMED_ENHMF;
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

            internal IStream GetStream()
            {
                if (medium.tymed == TYMED.TYMED_ISTREAM)
                {
                    var stream = (IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
                    return stream;
                }
                throw new InvalidComObjectException();
            }


            internal byte[] GetByteArray(int size = -1)
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

            internal uint GetDWORD()
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

            internal FILEDESCRIPTOR[] GetFileGroupDescriptor()
            {
                if (medium.tymed == TYMED.TYMED_HGLOBAL)
                {
                    long Length = ClipboardNative.GlobalSize(medium.unionmember).ToInt64();
                    IntPtr gLock = ClipboardNative.GlobalLock(medium.unionmember);
                    if (Length < Marshal.SizeOf<FILEGROUPDESCRIPTOR>()) { throw new ArgumentException(); }
                    try
                    {
                        FILEGROUPDESCRIPTOR fgd = Marshal.PtrToStructure<FILEGROUPDESCRIPTOR>(gLock);
                        var files = new FILEDESCRIPTOR[fgd.cItems];
                        for (int i = 0; i < fgd.cItems; i++)
                        {
                            files[i] = Marshal.PtrToStructure<FILEDESCRIPTOR>(gLock +
                                Marshal.SizeOf(fgd.cItems) +
                                Marshal.SizeOf<FILEDESCRIPTOR>() * i);
                        }
                        return files;
                    }
                    finally { ClipboardNative.GlobalUnlock(gLock); }
                }
                else
                {
                    throw new COMException();
                }
            }

            internal string GetString()
            {
                if (medium.tymed == TYMED.TYMED_HGLOBAL)
                {
                    byte[] bytes = GetByteArray();
                    switch (Format)
                    {
                        case "Text":
                        case "FileName":
                        case "OemText":
                        case "Rich Text Format":
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

            internal IEnumerable<string> GetDrop()
            {
                IntPtr Length = ClipboardNative.GlobalSize(medium.unionmember);
                IntPtr gLock = ClipboardNative.GlobalLock(medium.unionmember);
                try
                {
                    DROPFILES df = Marshal.PtrToStructure<DROPFILES>(gLock);
                    List<string> files = new List<string>();
                    string name;
                    int offset = (int)df.pFiles;
                    do
                    {
                        name = Marshal.PtrToStringAuto(gLock + offset);
                        files.Add(name);
                        offset += (name.Length + 1) * sizeof(char);
                    } while (name.Length != 0);
                    return files.ToArray();
                }
                finally
                {
                    ClipboardNative.GlobalUnlock(gLock);
                }
            }

            internal object GetMetafile(bool enhanced)
            {
                // IntPtr Length = ClipboardNative.GlobalSize(medium.unionmember);
                IntPtr hdc;
                hdc = Gdi32.GetDC(IntPtr.Zero);
                hdc = Gdi32.CreateCompatibleDC(Gdi32.GetDC(IntPtr.Zero));
                var PixelsX = (float)Gdi32.GetDeviceCaps(hdc, DeviceCap.HORZRES);
                var PixelsY = (float)Gdi32.GetDeviceCaps(hdc, DeviceCap.VERTRES);
                var MMX = (float)Gdi32.GetDeviceCaps(hdc, DeviceCap.HORZSIZE);
                var MMY = (float)Gdi32.GetDeviceCaps(hdc, DeviceCap.VERTSIZE);

                IntPtr hBitmap;
                int height;
                if (enhanced)
                {
                    ENHMETAHEADER header = new ENHMETAHEADER();
                    Gdi32.GetEnhMetaFileHeader(medium.unionmember, (uint)Marshal.SizeOf(header), ref header);
                    int width = header.rclBounds.Right - header.rclBounds.Left;
                    height = header.rclBounds.Bottom - header.rclBounds.Top;
                    hBitmap = CreateBitmap(width, height);
                    IntPtr old = Gdi32.SelectObject(hdc, hBitmap);
                    Gdi32.PlayEnhMetaFile(hdc, medium.unionmember, ref header.rclBounds);
                    Gdi32.SelectObject(hdc, old);
                }
                else
                {
                    IntPtr gLock = ClipboardNative.GlobalLock(medium.unionmember);
                    try
                    {
                        var metafilePict = Marshal.PtrToStructure<METAFILEPICT>(gLock);
                        Gdi32.SetMapMode(hdc, 8 /* MM_ANISOTROPIC */);
                        int width = metafilePict.xExt;
                        height = metafilePict.yExt;
                        hBitmap = CreateBitmap(width, height);
                        var bitmapOld = Gdi32.SelectObject(hdc, hBitmap);

                        Gdi32.SetGraphicsMode(hdc, 1 /* GM_ADVANCED */);
                        SIZE prevsize = new SIZE();
                        Gdi32.SetWindowExtEx(hdc, width, height, ref prevsize);
                        POINT prevorig = new POINT();
                        Gdi32.SetWindowOrgEx(hdc, 0, 0, ref prevorig);
                        Gdi32.SetViewportExtEx(hdc, width, height, ref prevsize);
                        Gdi32.PlayMetaFile(hdc, metafilePict.hMF);
                        Gdi32.SelectObject(hdc, bitmapOld);
                    }
                    finally
                    {
                        ClipboardNative.GlobalUnlock(gLock);
                    }
                }
                NativeBitmap.BitmapData bitmapData = new NativeBitmap.BitmapData();
                bitmapData.bitmapBytes = Gdi32.GetDIBits(hdc, hBitmap, (uint)height, out bitmapData.bitmapInfo);
                Gdi32.DeleteDC(hdc);
                return new MemoryStream(bitmapData.Bitmap);

                throw new NotImplementedException();
            }

            private static IntPtr CreateBitmap(int width, int height)
            {
                /* https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createcompatiblebitmap
                 * Note: When a memory device context is created, it initially has a 1-by-1 monochrome bitmap 
                 * selected into it. If this memory device context is used in CreateCompatibleBitmap, the 
                 * bitmap that is created is a monochrome bitmap. To create a color bitmap, use the HDC that 
                 * was used to create the memory device context
                 */
                return Gdi32.CreateCompatibleBitmap(Gdi32.GetDC(IntPtr.Zero), width, height);
            }
        }

        internal NativeStgMedium OleGetData(string format, int index = -1)
        {
            FORMATETC fmtEtc = new FORMATETC();
            fmtEtc.cfFormat = (short)formats[format];
            fmtEtc.tymed = GetPreferredMediumForFormat(format);
            fmtEtc.dwAspect = DVASPECT.DVASPECT_CONTENT; // DVASPECT_CONTENT
            fmtEtc.lindex = index;
            inner.GetData(ref fmtEtc, out STGMEDIUM medium);
            return new NativeStgMedium(medium, format);
        }
    }


}
