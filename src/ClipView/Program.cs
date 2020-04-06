using Clipboard;
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
        static int Main(string[] args)
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
*   Dump all formats
");
                    return 0;
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
            data.UseAscii = UseAscii;
            data.UseString = UseString;
            var formats = data.GetClipboardFormats();
            if (format == null)
            {
                Console.WriteLine(string.Join('\n', formats));
                return 0;
            }
            else if (format == "*")
            {
                foreach (var f in formats)
                {
                    Console.Write($"{f}: ");
                    Process(f, data, true);
                }
                return 0;
            }
            else
            {
                try
                {
                    return Process(format, data);
                }
                catch (KeyNotFoundException)
                {
                    Console.WriteLine($"Format {format} not found in clipboard. Available formats are: {string.Join(", ", data.GetClipboardFormats())}");
                    return -1;
                }
            }
        }

        private static int Process(string format, DataObject data, bool isStar = false)
        {
            var result = data.GetData(format);
            if (result is string)
            {
                Console.WriteLine(result);
            }
            else if (result is string[])
            {
                Console.WriteLine(string.Join('\n', result as string[]));
            }
            else if (result is FileContentsStream[])
            {
                foreach (var stream in (result as FileContentsStream[]))
                {
                    stream.Save(stream.FileName);
                    Console.WriteLine($"Saved file to {stream.FileName} ({stream.Length} bytes)");
                }
            }
            else if (result is FILEDESCRIPTOR[])
            {
                Console.WriteLine(string.Join(", ", (result as FILEDESCRIPTOR[])));
            }
            else if (result is MemoryStream && FormatIsBitmap(format))
            {
                SaveToFile(format, DibUtil.ImageFromClipboardDib(result as MemoryStream), "bmp", isStar);
            }
            else if (result is MemoryStream && format == "PNG")
            {
                SaveToFile(format, result as MemoryStream, "png", isStar);
            }
            else if (result is int)
            {
                Console.WriteLine(result);
            }
            else if (result is MemoryStream && WriteStream)
            {
                using (var file = File.Create($"clipboard_{format}.out"))
                {
                    (result as MemoryStream).WriteTo(file);
                    Console.WriteLine($"Saved raw stream to file clipboard_{format}.out");
                }
            }
            else
            {
                Console.WriteLine($"Couldn't infer data type {result.GetType().Name} for format {format}. Try dumping the contents as a string(-s, -a) or write the contents to a file (-o)");
                return -1;
            }
            return 0;
        }

        private static void SaveToFile(string format, MemoryStream result, string extension, bool appendFormat)
        {
            string formatSuffix = appendFormat ? $"_{format}" : "";
            string filename = extension != null ? $"clipboard{formatSuffix}.{extension}" : $"clipboard_{format}.out";
            using (var file = File.Create(filename))
            {
                result.WriteTo(file);
                Console.WriteLine($"Saved file to {filename}");
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

