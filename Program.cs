using System;
//using System.Windows.Forms;
// using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using clipview;
using System.Runtime.InteropServices.ComTypes;

namespace clipview
{
    public enum StandardClipboardFormats : uint
    {
        /// <summary>
        ///     Text format. Each line ends with a carriage return/linefeed (CR-LF) combination.A null character signals the end of
        ///     the data.Use this format for ANSI text.
        /// </summary>
        [Display(Name = "CF_TEXT")]
        Text = 1,

        /// <summary>
        ///     A handle to a bitmap (HBITMAP).
        /// </summary>
        [Display(Name = "CF_BITMAP")]
        Bitmap = 2,

        /// <summary>
        ///     Handle to a metafile picture format as defined by the METAFILEPICT structure.When passing a CF_METAFILEPICT handle
        ///     by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the
        ///     CF_METAFILEPICT handle.
        /// </summary>
        [Display(Name = "CF_METAFILEPICT")]
        MetafilePicture = 3,

        /// <summary>
        ///     Microsoft Symbolic Link (SYLK) format.
        /// </summary>
        [Display(Name = "CF_SYLK")]
        SymbolicLink = 4,

        /// <summary>
        ///     Software Arts' Data Interchange Format.
        /// </summary>
        [Display(Name = "CF_DIF")]
        DataInterchangeFormat = 5,

        /// <summary>
        ///     Tagged-image file format.
        /// </summary>
        [Display(Name = "CF_TIFF")]
        Tiff = 6,

        /// <summary>
        ///     Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF)
        ///     combination. A null character signals the end of the data.
        /// </summary>
        [Display(Name = "CF_OEMTEXT")]
        OemText = 7,

        /// <summary>
        ///     A memory object containing a BITMAPINFO structure followed by the bitmap bits.
        /// </summary>
        [Display(Name = "CF_DIB")]
        DeviceIndependentBitmap = 8,

        /// <summary>
        ///     Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color
        ///     palette, it should place the palette on the clipboard as well.
        ///     If the clipboard contains data in the CF_PALETTE (logical color palette) format, the application should use the
        ///     SelectPalette and RealizePalette functions to realize (compare) any other data in the clipboard against that
        ///     logical palette.
        ///     When displaying clipboard data, the clipboard always uses as its current palette any object on the clipboard that
        ///     is in the CF_PALETTE format.
        /// </summary>
        [Display(Name = "CF_PALETTE")]
        Palette = 9,

        /// <summary>
        ///     Data for the pen extensions to the Microsoft Windows for Pen Computing.
        /// </summary>
        [Display(Name = "CF_PENDATA")]
        PenData = 10,

        /// <summary>
        ///     Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
        /// </summary>
        [Display(Name = "CF_RIFF")]
        Riff = 11,

        /// <summary>
        ///     Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
        /// </summary>
        [Display(Name = "CF_WAVE")]
        Wave = 12,

        /// <summary>
        ///     Unicode text format.Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals
        ///     the end of the data.
        /// </summary>
        [Display(Name = "CF_UNICODETEXT")]
        UnicodeText = 13,

        /// <summary>
        ///     A handle to an enhanced metafile (HENHMETAFILE).
        /// </summary>
        [Display(Name = "CF_ENHMETAFILE")]
        EnhancedMetafile = 14,

        /// <summary>
        ///     A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by
        ///     passing the handle to the DragQueryFile function.
        /// </summary>
        [Display(Name = "CF_HDROP")]
        Drop = 15,

        /// <summary>
        ///     The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard,
        ///     if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the
        ///     current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text.
        ///     An application that pastes text from the clipboard can retrieve this format to determine which character set was
        ///     used to generate the text.
        ///     Note that the clipboard does not support plain text in multiple character sets.To achieve this, use a formatted
        ///     text data type such as RTF instead.
        ///     The system uses the code page associated with CF_LOCALE to implicitly convert from CF_TEXT to CF_UNICODETEXT.
        ///     Therefore, the correct code page table is used for the conversion.
        /// </summary>
        [Display(Name = "CF_LOCALE")]
        Locale = 16,

        /// <summary>
        ///     A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap
        ///     bits.
        /// </summary>
        [Display(Name = "CF_DIBV5")]
        DeviceIndependentBitmapV5 = 17,

        /// <summary>
        ///     Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the
        ///     WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The
        ///     hMem parameter must be NULL.
        /// </summary>
        [Display(Name = "CF_OWNERDISPLAY")]
        OwnerDisplay = 0x0080,

        /// <summary>
        ///     Text display format associated with a private format.
        ///     The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted
        ///     data.
        /// </summary>
        [Display(Name = "CF_DSPTEXT")]
        DisplayText = 0x0081,

        /// <summary>
        ///     Bitmap display format associated with a private format.
        ///     The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately
        ///     formatted data.
        /// </summary>
        [Display(Name = "CF_DSPBITMAP")]
        DisplayBitmap = 0x0082,

        /// <summary>
        ///     Metafile-picture display format associated with a private format.
        ///     The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the
        ///     privately formatted data.
        /// </summary>
        [Display(Name = "CF_DSPMETAFILEPICT")]
        DisplayMetafilePicture = 0x0083,

        /// <summary>
        ///     Enhanced metafile display format associated with a private format.
        ///     The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the
        ///     privately formatted data.
        /// </summary>
        [Display(Name = "CF_DSPENHMETAFILE")]
        DisplayEnhancedMetafile = 0x008E,

        /// <summary>
        ///     Start of a range of integer values for private clipboard formats.The range ends with CF_PRIVATELAST. Handles
        ///     associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles,
        ///     typically in response to the WM_DESTROYCLIPBOARD message.
        /// </summary>
        StartOfPrivateRange = 0x0200,

        /// <summary>
        ///     See CF_PRIVATEFIRST.
        /// </summary>
        EndOfPrivateRange = 0x02FF,

        /// <summary>
        ///     Start of a range of integer values for application-defined GDI object clipboard formats.The end of the range is
        ///     CF_GDIOBJLAST.
        ///     Handles associated with clipboard formats in this range are not automatically deleted using the GlobalFree function
        ///     when the clipboard is emptied.
        ///     Also, when using values in this range, the hMem parameter is not a handle to a GDI object, but is a handle
        ///     allocated by the GlobalAlloc function with the GMEM_MOVEABLE flag.
        /// </summary>
        StartOfApplicationDefinedGdiObjectRange = 0x0300,

        /// <summary>
        ///     See CF_GDIOBJFIRST.
        /// </summary>
        EndOfApplicationDefinedGdiObjectRange = 0x03FF
    }


    public class DataObject
    {
        private Dictionary<string, StandardClipboardFormats> formats = new Dictionary<string, StandardClipboardFormats>();

        private IDataObject inner;
        internal DataObject()
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
                string name = ClipboardHelper.ClipboardNative.GetClipboardFormatName((uint)fmt);
                RegisterFormat(name, (uint)fmt);
                // Console.WriteLine(fmt);
            }

            // for (uint i = 0, fmt = 0; i < ClipboardHelper.ClipboardNative.CountClipboardFormats(); i++)
            // {
            //     fmt = ClipboardHelper.ClipboardNative.EnumClipboardFormats(fmt);
            //     if (fmt != 0)
            //     {

            //     }
            // }
        }

        [DllImport("user32.dll")]
        private static extern uint RegisterClipboardFormat([MarshalAs(UnmanagedType.LPWStr)] string formatName);
        private void RegisterFormat(string name, uint fmt)
        {
            if (fmt < (uint)StandardClipboardFormats.StartOfPrivateRange)
            {
                formats[name] = (StandardClipboardFormats)fmt;
            } else if (!formats.ContainsKey(name)) {
                formats[name] = (StandardClipboardFormats) RegisterClipboardFormat(name);
            }
        }

        public override string ToString()
        {
            string[] textFormats = new string[] { "UnicodeText", "Text", "Locale", "OemText" };
            foreach (var f in textFormats)
            {
                if (formats.ContainsKey(f))
                {
                    byte[] data = InternalGetData(f);
                    Console.WriteLine(data.Length);
                    return Encoding.Unicode.GetString(data);
                }
            }
            return base.ToString();
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
            IntPtr ClipboardDataPointer = ClipboardHelper.ClipboardNative.GetClipboardData(fmt);

            //Do a bunch of crap necessary to copy the data from the memory
            //the above pointer points at to a place we can access it.
            IntPtr Length = ClipboardHelper.ClipboardNative.GlobalSize(ClipboardDataPointer);
            IntPtr gLock = ClipboardHelper.ClipboardNative.GlobalLock(ClipboardDataPointer);

            //Init a buffer which will contain the clipboard data
            byte[] Buffer = new byte[(int)Length];

            //Copy clipboard data to buffer
            Marshal.Copy(gLock, Buffer, 0, (int)Length);
            return Buffer;
        }

        public object GetData(string format)
        {
            return new MemoryStream(InternalGetData(format));
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
                    return TYMED.TYMED_HGLOBAL;
                case "FileContents":
                    return TYMED.TYMED_HGLOBAL;
            }
            return TYMED.TYMED_ISTREAM;
        }

        public class NativeStgMedium : IDisposable
        {
            STGMEDIUM medium;
            public string Format { get; private set; }
            public NativeStgMedium(STGMEDIUM medium, string format) { this.medium = medium; Format = format; }

            [DllImport("ole32.dll")]
            static extern void ReleaseStgMedium(STGMEDIUM medium);
            public void Dispose()
            {
                ReleaseStgMedium(medium);
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

            public byte[] GetByteArray()
            {
                if (medium.tymed != TYMED.TYMED_HGLOBAL) { throw new InvalidComObjectException(); }
                IntPtr Length = ClipboardHelper.ClipboardNative.GlobalSize(medium.unionmember);
                IntPtr gLock = ClipboardHelper.ClipboardNative.GlobalLock(medium.unionmember);

                //Init a buffer which will contain the clipboard data
                byte[] Buffer = new byte[(int)Length];

                //Copy clipboard data to buffer
                Marshal.Copy(gLock, Buffer, 0, (int)Length);
                return Buffer;
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
                            return Encoding.ASCII.GetString(bytes, 0, bytes.Length - 1);
                        case "UnicodeText":
                        case "FileNameW":
                            return Encoding.Unicode.GetString(bytes, 0, bytes.Length - 2);
                    }
                    return Encoding.Unicode.GetString(GetByteArray());
                }
                else if (medium.tymed == TYMED.TYMED_FILE)
                {
                    return Marshal.PtrToStringBSTR(medium.unionmember);
                }
                throw new InvalidComObjectException();
            }
        }

        public NativeStgMedium OleGetData(string format)
        {
            FORMATETC fmtEtc = new FORMATETC();
            fmtEtc.cfFormat = (short)formats[format];
            fmtEtc.tymed = GetPreferredMediumForFormat(format);
            fmtEtc.lindex = -1;
            inner.GetData(ref fmtEtc, out STGMEDIUM medium);
            return new NativeStgMedium(medium, format);
        }
    }
}


public class OleClipboardNative
{
    [DllImport("ole32.dll", PreserveSig = false)]
    // [return: MarshalAs(UnmanagedType.IUnknown)]
    public static extern IDataObject OleGetClipboard();

    // [StructLayout(LayoutKind.Sequential)]
    // public struct STGMEDIUM
    // {
    //     [MarshalAs(UnmanagedType.U4)]
    //     public int tymed;
    //     public IntPtr data;
    //     [MarshalAs(UnmanagedType.IUnknown)]
    //     public object pUnkForRelease;
    // }


    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000103-0000-0000-C000-000000000046")]
    public interface IEnumFORMATETC
    {
        [PreserveSig]
        int Next(int celt, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] FORMATETC[] rgelt, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pceltFetched);
        [PreserveSig]
        int Skip(int celt);
        [PreserveSig]
        int Reset();
        void Clone(out IEnumFORMATETC newEnum);
    }

    // public enum DATADIR
    // {
    //     Get = 1,
    //     Set = 2
    // }

    /*
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000010E-0000-0000-C000-000000000046")]
        public interface IDataObject
        {
            void GetData([In] ref FORMATETC format, out STGMEDIUM medium);
            void GetDataHere([In] ref FORMATETC format, ref STGMEDIUM medium);
            [PreserveSig]
            int QueryGetData([In] ref FORMATETC format);
            [PreserveSig]
            int GetCanonicalFormatEtc([In] ref FORMATETC formatIn, out FORMATETC formatOut);
            void SetData([In] ref FORMATETC formatIn, [In] ref STGMEDIUM medium, [MarshalAs(UnmanagedType.Bool)] bool release);
            IEnumFORMATETC EnumFormatEtc(DATADIR direction);

            [PreserveSig]
            int DAdvise([In] ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection);
            void DUnadvise(int connection);
            [PreserveSig]
            int EnumDAdvise(out IEnumSTATDATA enumAdvise);
        }
    */
    // [StructLayout(LayoutKind.Sequential)]
    // public struct FORMATETC
    // {
    //     public short cfFormat;
    //     public IntPtr ptd;
    //     [MarshalAs(UnmanagedType.U4)]
    //     public DVASPECT dwAspect;
    //     public int lindex;
    //     [MarshalAs(UnmanagedType.U4)]
    //     public TYMED tymed;
    // };

    ///// <summary>
    ///// The DVASPECT enumeration values specify the desired data or view aspect of the object when drawing or getting data.
    ///// </summary>
    // [Flags]
    // public enum DVASPECT
    // {
    //     DVASPECT_CONTENT = 1,
    //     DVASPECT_THUMBNAIL = 2,
    //     DVASPECT_ICON = 4,
    //     DVASPECT_DOCPRINT = 8
    // }

    // Summary:
    //     Provides the managed definition of the TYMED structure.
    // [Flags]
    // public enum TYMED
    // {
    //     // Summary:
    //     //     No data is being passed.
    //     TYMED_NULL = 0,
    //     //
    //     // Summary:
    //     //     The storage medium is a global memory handle (HGLOBAL). Allocate the global
    //     //     handle with the GMEM_SHARE flag. If the System.Runtime.InteropServices.ComTypes.STGMEDIUMSystem.Runtime.InteropServices.ComTypes.STGMEDIUM.pUnkForRelease
    //     //     member is null, the destination process should use GlobalFree to release
    //     //     the memory.
    //     TYMED_HGLOBAL = 1,
    //     //
    //     // Summary:
    //     //     The storage medium is a disk file identified by a path. If the STGMEDIUMSystem.Runtime.InteropServices.ComTypes.STGMEDIUM.pUnkForRelease
    //     //     member is null, the destination process should use OpenFile to delete the
    //     //     file.
    //     TYMED_FILE = 2,
    //     //
    //     // Summary:
    //     //     The storage medium is a stream object identified by an IStream pointer. Use
    //     //     ISequentialStream::Read to read the data. If the System.Runtime.InteropServices.ComTypes.STGMEDIUMSystem.Runtime.InteropServices.ComTypes.STGMEDIUM.pUnkForRelease
    //     //     member is not null, the destination process should use IStream::Release to
    //     //     release the stream component.
    //     TYMED_ISTREAM = 4,
    //     //
    //     // Summary:
    //     //     The storage medium is a storage component identified by an IStorage pointer.
    //     //     The data is in the streams and storages contained by this IStorage instance.
    //     //     If the System.Runtime.InteropServices.ComTypes.STGMEDIUMSystem.Runtime.InteropServices.ComTypes.STGMEDIUM.pUnkForRelease
    //     //     member is not null, the destination process should use IStorage::Release
    //     //     to release the storage component.
    //     TYMED_ISTORAGE = 8,
    //     //
    //     // Summary:
    //     //     The storage medium is a Graphics Device Interface (GDI) component (HBITMAP).
    //     //     If the System.Runtime.InteropServices.ComTypes.STGMEDIUMSystem.Runtime.InteropServices.ComTypes.STGMEDIUM.pUnkForRelease
    //     //     member is null, the destination process should use DeleteObject to delete
    //     //     the bitmap.
    //     TYMED_GDI = 16,
    //     //
    //     // Summary:
    //     //     The storage medium is a metafile (HMETAFILE). Use the Windows or WIN32 functions
    //     //     to access the metafile's data. If the System.Runtime.InteropServices.ComTypes.STGMEDIUMSystem.Runtime.InteropServices.ComTypes.STGMEDIUM.pUnkForRelease
    //     //     member is null, the destination process should use DeleteMetaFile to delete
    //     //     the bitmap.
    //     TYMED_MFPICT = 32,
    //     //
    //     // Summary:
    //     //     The storage medium is an enhanced metafile. If the System.Runtime.InteropServices.ComTypes.STGMEDIUMSystem.Runtime.InteropServices.ComTypes.STGMEDIUM.pUnkForRelease
    //     //     member is null, the destination process should use DeleteEnhMetaFile to delete
    //     //     the bitmap.
    //     TYMED_ENHMF = 64,
    // }

}

public class ClipboardHelper : IDisposable
{
    private bool result;
    public ClipboardHelper()
    {
        result = ClipboardNative.OpenClipboard(IntPtr.Zero);
    }
    public void Dispose()
    {
        if (result)
        {
            ClipboardNative.CloseClipboard();
        }
    }

    public DataObject GetDataObject()
    {
        return new DataObject();
    }

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
        static extern bool GlobalUnlock(IntPtr hMem);

    }

}


static class DibUtil
{
    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    struct BITMAPFILEHEADER
    {
        public static readonly short BM = 0x4d42; // BM

        public short bfType;
        public int bfSize;
        public short bfReserved1;
        public short bfReserved2;
        public int bfOffBits;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct BITMAPINFOHEADER
    {
        public int biSize;
        public int biWidth;
        public int biHeight;
        public short biPlanes;
        public short biBitCount;
        public int biCompression;
        public int biSizeImage;
        public int biXPelsPerMeter;
        public int biYPelsPerMeter;
        public int biClrUsed;
        public int biClrImportant;
    }
    public static class BinaryStructConverter
    {
        public static T FromByteArray<T>(byte[] bytes) where T : struct
        {
            IntPtr ptr = IntPtr.Zero;
            try
            {
                int size = Marshal.SizeOf(typeof(T));
                ptr = Marshal.AllocHGlobal(size);
                Marshal.Copy(bytes, 0, ptr, size);
                object obj = Marshal.PtrToStructure(ptr, typeof(T));
                return (T)obj;
            }
            finally
            {
                if (ptr != IntPtr.Zero)
                    Marshal.FreeHGlobal(ptr);
            }
        }

        public static byte[] ToByteArray<T>(T obj) where T : struct
        {
            IntPtr ptr = IntPtr.Zero;
            try
            {
                int size = Marshal.SizeOf(typeof(T));
                ptr = Marshal.AllocHGlobal(size);
                Marshal.StructureToPtr(obj, ptr, true);
                byte[] bytes = new byte[size];
                Marshal.Copy(ptr, bytes, 0, size);
                return bytes;
            }
            finally
            {
                if (ptr != IntPtr.Zero)
                    Marshal.FreeHGlobal(ptr);
            }
        }
    }

    public static MemoryStream ImageFromClipboardDib(MemoryStream ms)
    {
        if (ms != null)
        {
            byte[] dibBuffer = new byte[ms.Length];
            ms.Read(dibBuffer, 0, dibBuffer.Length);

            BITMAPINFOHEADER infoHeader =
                BinaryStructConverter.FromByteArray<BITMAPINFOHEADER>(dibBuffer);

            int fileHeaderSize = Marshal.SizeOf(typeof(BITMAPFILEHEADER));
            int infoHeaderSize = infoHeader.biSize;
            int fileSize = fileHeaderSize + infoHeader.biSize + infoHeader.biSizeImage;

            BITMAPFILEHEADER fileHeader = new BITMAPFILEHEADER();
            fileHeader.bfType = BITMAPFILEHEADER.BM;
            fileHeader.bfSize = fileSize;
            fileHeader.bfReserved1 = 0;
            fileHeader.bfReserved2 = 0;
            fileHeader.bfOffBits = fileHeaderSize + infoHeaderSize + infoHeader.biClrUsed * 4;

            byte[] fileHeaderBytes =
                BinaryStructConverter.ToByteArray<BITMAPFILEHEADER>(fileHeader);

            MemoryStream msBitmap = new MemoryStream();
            msBitmap.Write(fileHeaderBytes, 0, fileHeaderSize);
            msBitmap.Write(dibBuffer, 0, dibBuffer.Length);
            msBitmap.Seek(0, SeekOrigin.Begin);
            return msBitmap;
        }
        else
        {
            throw new ArgumentException("No source data provided");
        }
    }

}

class Program
{
    private static void SaveToFile(string filename, MemoryStream input)
    {
        using (var fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write))
        {
            input.WriteTo(fileStream);
            fileStream.Flush();
        }
        Console.WriteLine($"Saved to {filename}");
    }


    /*
                RenPrivateSourceFolder,
                RenPrivateMessages, 
                RenPrivateItem, 
                FileGroupDescriptor, 
                FileGroupDescriptorW, 
                FileDrop, 
                FileNameW, 
                FileName, 
                FileContents, 
                Object_Descriptor, 
                System_String, 
                UnicodeText, 
                Text, 
                CSV

    */
    Dictionary<string, Action<object, string>> knownFormats = new Dictionary<string, Action<object, string>>{
            // {"Text", (object data, string filename) => {}},
            {"DeviceIndependentBitmap", (object data, string filename) => {
                MemoryStream imgStream = DibUtil.ImageFromClipboardDib(data as MemoryStream);
                if (imgStream == null) { throw new Exception("Couldn't create image from clipboard content"); }
                SaveToFile(filename + ".bmp", imgStream);
            }},
            {"PNG", (object data, string filename) => { SaveToFile(filename + ".png", data as MemoryStream); }},
        };


    private void GetFormat(DataObject content, string format)
    {
        object data = content.GetData(format);
        Console.WriteLine("data type = " + data.GetType().Name);
        const string filename = "clipboard";
        if (knownFormats.ContainsKey(format))
        {
            knownFormats[format](data, filename);
        }
        // else if (data is Bitmap)
        // {
        //     Bitmap b = (Bitmap)data;
        //     b.Save(filename + ".bmp");
        //     Console.WriteLine($"Saved to {filename}.bmp");
        // }
        else if (data is MemoryStream)
        {
            MemoryStream ms = data as MemoryStream;
            // Is it a scalar?
            if (ms.Length <= 8)
            {
                string value = new StreamReader(ms).ReadToEnd();
                if (int.TryParse(value, out int intvalue))
                {
                    Console.WriteLine($"Integer value: {intvalue}");
                    return;
                }
                else if (bool.TryParse(value, out bool boolvalue))
                {
                    Console.WriteLine($"Bool value: {boolvalue}");
                    return;
                }
                else if (ms.Length == 4)
                {
                    int x = BitConverter.ToInt32(Encoding.UTF8.GetBytes(value));
                    Console.WriteLine($"Int value: {x}");
                    return;
                }
            }
            SaveToFile(filename + ".out", ms);

        }
        else if (data is string || data is string[])
        {
            if (data is string[])
            {
                data = string.Join(Environment.NewLine, data as string[]) as string;
            }
            if (PrintOut)
            {
                Console.WriteLine((string)data);
            }
            else
            {
                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes((string)data));
                SaveToFile(filename + ".txt", ms);
            }
        }
        else if (data is File)
        {
            Console.WriteLine("Found a file");
        }
        else if (data != null)
        {
            Console.Error.WriteLine($"Found format {format} but handling not yet implemented. Type is {data.GetType().Name}");
            return;
        }
        else
        {
            Console.Error.WriteLine($"Found format {format} but GetData returned null");
            return;
        }
    }

    private bool PrintOut { get; set; }
    [STAThread]
    static void Main(string[] args)
    {
        string format = args.Length > 0 ? args[0] : null;
        // using (var clip = new ClipboardHelper())
        // {
        //     DataObject content = clip.GetDataObject();
        //     var formats = content.GetClipboardFormats();
        //     Console.WriteLine(string.Join(' ', formats));
        // }


        var data = new DataObject();
        if (format == null)
        {
            var formats = data.GetClipboardFormats();
            Console.WriteLine(string.Join(',', formats));
        }
        else
        {
            using (var medium = data.OleGetData(format))
            {
                // MemoryStream ms = new MemoryStream();
                // medium.GetStream().CopyTo()
                string r = medium.GetString();
                Console.WriteLine(r);
            }
        }

        /*
                using (var clip = new ClipboardHelper())
                {
                    DataObject content = clip.GetDataObject();
                    var formats = content.GetClipboardFormats();
                    Program p = new Program();

                    if (args.Length >= 2 && args[1].ToUpper() == "-P")
                    {
                        p.PrintOut = true;
                    }
                    if (format != null)
                    {
                        if (formats.Contains(format))
                        {
                            p.GetFormat(content, format);
                        }
                        else
                        {
                            Console.Error.WriteLine($"Format {format} not available");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Available formats: {string.Join(", ", formats)}");
                        Console.WriteLine(content);
                    }
                }
                */


    }
}
