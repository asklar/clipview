using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace clipview
{
    public class OleClipboardNative
    {
        [DllImport("ole32.dll", PreserveSig = false)]
        public static extern void OleInitialize(IntPtr intPtr);

        [DllImport("ole32.dll", PreserveSig = false)]
        public static extern IDataObject OleGetClipboard();
    }

    public class ClipboardHelper : IDisposable
    {
        private readonly bool result;
        public ClipboardHelper()
        {
            result = ClipboardNative.OpenClipboard(IntPtr.Zero);
        }
        public void Dispose()
        {
            if (result)
            {
                ClipboardNative.CloseClipboard();
            }
        }

        public DataObject GetDataObject()
        {
            return new DataObject();
        }



    }
}