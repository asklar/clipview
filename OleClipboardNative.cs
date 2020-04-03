using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace clipview
{
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



    }
}