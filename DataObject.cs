using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Linq;

namespace clipview
{
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
                string name = ClipboardNative.GetClipboardFormatName((uint)fmt);
                RegisterFormat(name, (uint)fmt);
                // Console.WriteLine(fmt);
            }

            // for (uint i = 0, fmt = 0; i < ClipboardNative.CountClipboardFormats(); i++)
            // {
            //     fmt = ClipboardNative.EnumClipboardFormats(fmt);
            //     if (fmt != 0)
            //     {

            //     }
            // }
        }

        [DllImport("user32.dll", SetLastError=true)]
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
                Console.WriteLine($"Registering format {name} {formats[name]} {value}");
            }
        }

        public override string ToString()
        {
            string[] textFormats = new string[] { "UnicodeText", "Text", "Locale", "OemText" };
            foreach (var f in textFormats)
            {
                if (formats.ContainsKey(f))
                {
                    using (var medium = OleGetData(f))
                    {
                        return medium.GetString();
                    }
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
                    return TYMED.TYMED_ISTREAM;
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

            public byte[] GetByteArray(int size = -1)
            {
                if (medium.tymed != TYMED.TYMED_HGLOBAL) { throw new InvalidComObjectException(); }
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
                    return Encoding.Unicode.GetString(GetByteArray());
                }
                else if (medium.tymed == TYMED.TYMED_FILE)
                {
                    return Marshal.PtrToStringBSTR(medium.unionmember);
                }
                throw new InvalidComObjectException();
            }
        }

        internal NativeStgMedium OleGetData(string format)
        {
            FORMATETC fmtEtc = new FORMATETC();
            fmtEtc.cfFormat = (short)formats[format];
            fmtEtc.tymed = GetPreferredMediumForFormat(format);
            //fmtEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
            fmtEtc.lindex = -1;
            inner.GetData(ref fmtEtc, out STGMEDIUM medium);
            return new NativeStgMedium(medium, format);
        }
    }

}