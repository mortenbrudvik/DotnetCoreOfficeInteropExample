using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using OfficeInteropLib.ComInterop;
using Excel = Microsoft.Office.Interop.Excel; 

namespace OfficeInteropLib
{
    /// <summary>
    /// Make sure to use the using statement to ensure cleanup of com objects.
    /// </summary>
    public class ExcelApi : IDisposable
    {
        private readonly Excel.Application _excel;
        private readonly Excel.Workbooks _workbooks;
        private readonly Excel.Windows _windows;

        // NB! will throw an exception if there is no running excel instances
        public ExcelApi()
        {
            _excel = (Excel.Application) MarshalCore.GetActiveObject("Excel.Application");
            _workbooks = (Excel.Workbooks) _excel.Workbooks;
            _windows = (Excel.Windows) _excel.Windows;
        }

        public string Version => _excel.Version;

        public IEnumerable<WindowInfo> GetWindows()
        {
            if (_excel == null)
                yield break;

            for (var i = 1; i <= _workbooks.Count; i++)
            {
                Excel.Workbook workbook = _workbooks[i];

                var windowInfo = GetWindow(workbook);

                Marshal.ReleaseComObject(workbook);

                yield return windowInfo;
            }
        }

        // Alternative way of fetching the workbooks using Running object table (Seem to be more stable)
        public static IEnumerable<WindowInfo> GetWindows2()
        {
            var runningObjects = RunningObjectTable.GetObjects();
            foreach (var obj in runningObjects)
            {
                if (obj is Excel.Workbook workbook)
                {
                    var windowInfo = GetWindow(workbook);

                    Marshal.ReleaseComObject(workbook);

                    yield return windowInfo;
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
                if (comObject is Excel.Workbook workbook)
                    exist = true;
                Marshal.ReleaseComObject(comObject);
            }

            return exist;
        }

        private static WindowInfo GetWindow(Excel.Workbook workbook)
        {
            Excel.Windows workbookWindows = workbook.Windows;
            var windowHandles = new List<int>();
            for (var j = 1; j <= workbookWindows.Count; j++)
            {
                Excel.Window workbookWindow = workbookWindows[j];

                windowHandles.Add(workbookWindow.Hwnd);

                Marshal.ReleaseComObject(workbookWindow);
            }

            var path = workbook.FullName;
            Marshal.ReleaseComObject(workbookWindows);

            return new WindowInfo(new IntPtr(windowHandles.SingleOrDefault()), path);
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(_workbooks);
            Marshal.ReleaseComObject(_windows);
            Marshal.ReleaseComObject(_excel);
        }
    }
}