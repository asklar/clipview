using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace clipview
{
    public class FileContentsStream : IDisposable
    {
        private STATSTG stat;
        private IStream stream { get; set; }
        public FileContentsStream(IStream stream)
        {
            this.stream = stream;
            stream.Stat(out stat, 0);
        }
        public string FileName { get => stat.pwcsName; }
        public long Length { get => stat.cbSize; }
        enum STREAM_SEEK : int
        {
            STREAM_SEEK_SET,
            STREAM_SEEK_CUR,
            STREAM_SEEK_END
        };

        public void SaveToStream(Stream outputStream)
        {
            stream.Seek(0, (int)STREAM_SEEK.STREAM_SEEK_SET, IntPtr.Zero);
            byte[] buffer = new byte[512];
            int cbRead = 0;
            unsafe
            {
                IntPtr pcbRead = new IntPtr((void*)&cbRead);
                try
                {
                    do
                    {
                        stream.Read(buffer, buffer.Length, pcbRead);
                        outputStream.Write(buffer, 0, cbRead);
                    } while (cbRead >= buffer.Length);
                }
                catch (EndOfStreamException) { return; }
            }
        }

        public void Save(string filepath)
        {
            using (var file = File.Create(filepath))
            {
                SaveToStream(file);
            }
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(stream);
        }
    }

}