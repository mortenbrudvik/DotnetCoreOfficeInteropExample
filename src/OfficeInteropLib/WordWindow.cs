using System;

namespace OfficeInteropLib
{
    public class WordWindow
    {
        public WordWindow(IntPtr handle, string documentPath)
        {
            Handle = handle;
            DocumentPath = documentPath;

        }

        public IntPtr Handle { get; }
        public string DocumentPath { get; }
    };
}