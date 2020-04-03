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
            try {
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
            // } catch (Exception e) {
            //     Console.WriteLine(e);
            }
            finally{}
        }
    }
}

