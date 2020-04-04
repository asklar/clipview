using System;
using System.IO;

namespace clipview
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            OleClipboardNative.OleInitialize(IntPtr.Zero);

            string format = args.Length > 0 ? args[0] : null;

            var data = new DataObject();
            if (format == null)
            {
                var formats = data.GetClipboardFormats();
                Console.WriteLine(string.Join(',', formats));
            }
            else
            {
                var result = data.GetData(format);
                if (result is string)
                {
                    Console.WriteLine(result);
                }
                else if (result is FileContentsStream[])
                {
                    foreach (var stream in (result as FileContentsStream[]))
                    {
                        stream.Save(stream.FileName);
                        Console.WriteLine($"Saved file to {stream.FileName} ({stream.Length} bytes)");
                    }
                }
                else if (result is MemoryStream && FormatIsBitmap(format))
                {
                    using (var file = File.Create("clipboard.bmp"))
                    {
                        DibUtil.ImageFromClipboardDib(result as MemoryStream).WriteTo(file);
                        Console.WriteLine($"Saved file to clipboard.bmp");
                    }
                }
                else
                {
                    Console.WriteLine(result.GetType().Name);
                }
            }
        }

        private static bool FormatIsBitmap(string format)
        {
            switch (format)
            {
                case "Bitmap":
                case "DeviceIndependentBitmap":
                case "DeviceIndependentBitmapW":
                case "DeviceIndependentBitmapV5":
                    return true;
                default:
                    return false;
            }
        }
    }
}

