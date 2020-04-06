using System;
using System.Runtime.InteropServices;

namespace Clipboard
{
    internal class NativeBitmap : IDisposable
    {
        private readonly IntPtr hBitmap;
        private BITMAP bitmap;
        public NativeBitmap(IntPtr hBitmap)
        {
            this.hBitmap = hBitmap;
            bitmap = new BITMAP();
            Gdi32.GetObjectBitmap(hBitmap, Marshal.SizeOf(bitmap), ref bitmap);
            var hdc = Gdi32.GetDC(IntPtr.Zero);

            bitmapData.bitmapBytes = Gdi32.GetDIBits(hdc, hBitmap, (uint)bitmap.bmHeight, out bitmapData.bitmapInfo);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BitmapData
        {
            public BITMAPINFO bitmapInfo;
            public byte[] bitmapBytes;
            public byte[] Bitmap
            {
                get
                {
                    byte[] ret = new byte[Marshal.SizeOf(bitmapInfo) + bitmapBytes.Length];
                    byte[] info = BinaryStructConverter.ToByteArray(bitmapInfo);
                    info.CopyTo(ret, 0);
                    bitmapBytes.CopyTo(ret, info.Length);
                    return ret;
                }
            }
        };

        private BitmapData bitmapData = new BitmapData();
        public byte[] Bitmap => bitmapData.Bitmap;
        public void Dispose()
        {
            Gdi32.DeleteObject(hBitmap);
            Marshal.FreeHGlobal(bitmap.bmBits);
        }

    }
}