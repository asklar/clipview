using System;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

namespace clipview
{
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
        static Dictionary<string, Action<object, string>> knownFormats = new Dictionary<string, Action<object, string>>{
            // {"Text", (object data, string filename) => {}},
            {"DeviceIndependentBitmap", (object data, string filename) => {
                MemoryStream imgStream = DibUtil.ImageFromClipboardDib(data as MemoryStream);
                if (imgStream == null) { throw new Exception("Couldn't create image from clipboard content"); }
                SaveToFile(filename + ".bmp", imgStream);
            }},
        };

        [STAThread]
        static void Main(string[] args)
        {
            string format = args.Length > 0 ? args[0] : null;
            IDataObject content = Clipboard.GetDataObject();
            string[] formats = content.GetFormats();
            Console.Out.Flush();
            if (format != null)
            {
                if (formats.Contains(format))
                {
                    object data = content.GetData(format);
                    const string filename = "clipboard";
                    if (knownFormats.ContainsKey(format))
                    {
                        knownFormats[format](data, filename);
                    }
                    else if (data is MemoryStream)
                    {
                        SaveToFile(filename + ".out", data as MemoryStream);
                    }
                    else if (data is string || data is string[])
                    {
                        if (data is string[]) {
                            data = string.Join(Environment.NewLine, data as string[]) as string;
                        }
                        if (args.Length >= 2 && args[1].ToUpper() == "-P")
                        {
                            Console.WriteLine((string)data);
                        }
                        else
                        {
                            MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes((string)data));
                            SaveToFile(filename + ".txt", ms);
                        }
                    } else if (data is File) {
                        Console.WriteLine("Found a file");
                    }
                    else if (data != null)
                    {
                        Console.Error.WriteLine($"Found format {format} but handling not yet implemented. Type is {data.GetType().Name}");
                        return;
                    }
                    else {
                        Console.Error.WriteLine($"Found format {format} but GetData returned null");
                        return;
                    }
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
    }
}
