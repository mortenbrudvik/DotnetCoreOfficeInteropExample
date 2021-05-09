using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OfficeInteropLib.ComInterop;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeInteropLib
{
    /// <summary>
    /// Make sure to use the using statement to ensure cleanup of com objects.
    /// </summary>
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

        public IEnumerable<WindowInfo> GetWindows()
        {
            for (var i = 1; i <= _windows.Count; i++)
            {
                Word.Window window =  _windows[i];

                var handle = new IntPtr(window.Hwnd);
                var path = window.Document.FullName;
                Marshal.ReleaseComObject(window);
                yield return new WindowInfo(handle, path);
            }
        }

        public static IEnumerable<WindowInfo> GetWindows2()
        {
            var runningObjects = RunningObjectTable.GetObjects();
            foreach (var obj in runningObjects)
            {
                if (obj is Word.Application application)
                {
                    Word.Windows windows = application.Windows;
                    for (var i = 1; i <= windows.Count; i++)
                    {
                        Word.Window window =  windows[i];

                        var handle = new IntPtr(window.Hwnd);
                        var path = window.Document.FullName;
                        Marshal.ReleaseComObject(window);
                        yield return new WindowInfo(handle, path);
                    }

                    Marshal.ReleaseComObject(windows);
                }

                Marshal.ReleaseComObject(obj);
            }
        }

        public static bool HasRunningInstances()
        {
            var exist = false;
            var comObjects = RunningObjectTable.GetObjects();
            foreach (object comObject in comObjects)
            {
                if (comObject is Word.Window window)
                    exist = true;
                Marshal.ReleaseComObject(comObject);
            }

            return exist;
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(_windows);
            Marshal.ReleaseComObject(_word);
        }
    }
}