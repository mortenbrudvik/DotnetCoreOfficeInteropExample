using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OfficeInteropLib.Common;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeInteropLib
{
    public class WordApi : IDisposable
    {
        private readonly Word.Application _word;
        private readonly Word.Windows _windows;

        public WordApi()
        {
            _word = (Word.Application)MarshalCore.GetActiveObject("Word.Application");
            _windows = _word.Windows;
        }

        public string Version => _word.Version;

        public IEnumerable<WordWindow> GetWindows()
        {
            for (var i = 1; i <= _windows.Count; i++)
            {
                Word.Window window =  _windows[i];

                var handle = new IntPtr(window.Hwnd);
                var path = window.Document.FullName;
                Marshal.ReleaseComObject(window);
                yield return new WordWindow(handle, path);
            }
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(_windows);
            Marshal.ReleaseComObject(_word);
        }
    }
}