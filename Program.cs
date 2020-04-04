using System;
using System.Collections.Generic;
using System.IO;

namespace clipview
{
    class Program
    {
        public static bool UseAscii { get; set; } = false;
        public static bool UseString { get; set; } = false;
        public static bool WriteStream { get; set; } = false;
        [STAThread]
        static void Main(string[] args)
        {
            OleClipboardNative.OleInitialize(IntPtr.Zero);

            string format = null;
            foreach (var arg in args)
            {
                if (arg == "-a")
                {
                    UseAscii = true;
                }
                else if (arg == "-s")
                {
                    UseString = true;
                }
                else if (arg == "-o")
                {
                    WriteStream = true;
                }
                else if (arg == "-?")
                {
                    WriteConsole("ClipView", ConsoleColor.White, ConsoleColor.Black);
                    Console.WriteLine(" -- dump clipboard contents.");
                    WriteConsole("https://github.com/asklar/clipview", ConsoleColor.Blue, ConsoleColor.Gray);
                    Console.WriteLine(@"

Options:
-s  Assume strings for unrecognized formats (default is stream)
-a  Assume ASCII for unrecognized formats (default is Unicode)
-o  Dump stream to file for unrecognized formats
");
                    return;
                }
                else if (format == null)
                {
                    format = arg;
                }
                else
                {
                    throw new ArgumentException(arg);
                }
            }

            var data = new DataObject();
            if (format == null)
            {
                var formats = data.GetClipboardFormats();
                Console.WriteLine(string.Join('\n', formats));
            }
            else
            {
                try
                {
                    Process(format, data);
                }
                catch (KeyNotFoundException)
                {
                    Console.WriteLine($"Format {format} not found in clipboard. Available formats are: {string.Join(", ", data.GetClipboardFormats())}");
                    return;
                }
            }
        }

        private static void Process(string format, DataObject data)
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
            else if (result is MemoryStream && format == "PNG")
            {
                throw new NotImplementedException("PNG");
            }
            else if (result is MemoryStream && WriteStream)
            {
                using (var file = File.Create("clipboard.out"))
                {
                    (result as MemoryStream).WriteTo(file);
                }
            }
            else
            {
                Console.WriteLine(result.GetType().Name);
            }
        }

        private static void WriteConsole(string s, ConsoleColor f, ConsoleColor b)
        {
            var oldF = Console.ForegroundColor;
            var oldB = Console.BackgroundColor;
            Console.ForegroundColor = f;
            Console.BackgroundColor = b;
            Console.Write(s);
            Console.ForegroundColor = oldF;
            Console.BackgroundColor = oldB;
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

