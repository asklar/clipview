﻿using Clipboard;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace clipview
{
    class Program
    {
        public bool UseAscii { get; set; } = false;
        public bool UseString { get; set; } = false;
        public bool WriteStream { get; set; } = false;

        private string ParseOptions()
        {
            string format = null;
            var args = Environment.GetCommandLineArgs().Skip(1);
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
                    Environment.Exit(0);
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

            return format;
        }

        [STAThread]
        static int Main(string[] args)
        {
            OleClipboardNative.OleInitialize(IntPtr.Zero);
            Program p = new Program();
            return p.Do();
        }

        private Program() { Format = ParseOptions(); }
        private string Format { get; set; }
        private int Do()
        {
            var data = new DataObject() { UseAscii = UseAscii, UseString = UseString };
            var formats = data.GetClipboardFormats();
            if (Format == null)
            {
                Console.WriteLine(string.Join('\n', formats));
                return 0;
            }
            else if (Format == "*")
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
                    return Process(Format, data);
                }
                catch (KeyNotFoundException)
                {
                    Console.WriteLine($"Format {Format} not found in clipboard. Available formats are: {string.Join(", ", data.GetClipboardFormats())}");
                    return -1;
                }
                catch (NotImplementedException)
                {
                    Console.WriteLine($"Support for format {Format} is not yet implemented");
                    return -2;
                }
            }
        }

        private int Process(string format, DataObject data, bool isStar = false)
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
            else if (result is MemoryStream && FormatIsStreamToFile(format))
            {
                SaveToFile(format, result as MemoryStream, format.ToLowerInvariant(), isStar);
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

        private static bool FormatIsStreamToFile(string format)
        {
            return format == "PNG" || format == "GIF" || format == "JFIF";
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
                case "MetafilePicture":
                case "EnhancedMetafile":
                    return true;
                default:
                    return false;
            }
        }
    }
}

