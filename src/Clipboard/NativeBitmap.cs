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
        };
        private BitmapData bitmapData = new BitmapData();

        public byte[] Bitmap
        {
            get
            {
                byte[] ret = new byte[Marshal.SizeOf(bitmapData.bitmapInfo) + bitmapData.bitmapBytes.Length];
                byte[] info = BinaryStructConverter.ToByteArray(bitmapData.bitmapInfo);
                info.CopyTo(ret, 0);
                bitmapData.bitmapBytes.CopyTo(ret, info.Length);
                return ret;
            }
        }
        public void Dispose()
        {
            Gdi32.DeleteObject(hBitmap);
            Marshal.FreeHGlobal(bitmap.bmBits);
        }

    }
}