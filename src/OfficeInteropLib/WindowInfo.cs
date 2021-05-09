using System;

namespace OfficeInteropLib
{
    public class WindowInfo
    {
        public WindowInfo(IntPtr handle, string documentPath)
        {
            Handle = handle;
            DocumentPath = documentPath;
        }

        public IntPtr Handle { get; }
        public string DocumentPath { get; }

        public override string ToString() => $"Handle: {Handle}, Path: {DocumentPath}";
    };
}