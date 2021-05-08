using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OfficeInteropLib.Common;
using Excel = Microsoft.Office.Interop.Excel; 

namespace OfficeInteropLib
{
    public class ExcelApi : IDisposable
    {
        private readonly Excel.Application _excel;
        private readonly Excel.Workbooks  _workbooks;
        private readonly Excel.Windows  _windows;

        public ExcelApi()
        {
            _excel = (Excel.Application)MarshalCore.GetActiveObject("Excel.Application");
            _workbooks = (Excel.Workbooks)_excel.Workbooks;
            _windows = (Excel.Windows)_excel.Windows;
        }

        public string Version => _excel.Version;

        public IEnumerable<WordWindow> GetWindows()
        {
            var winCount = _windows.Count;
            
            for (var i = 1; i <= _workbooks.Count; i++)
            {
                Excel.Workbook workbook = _workbooks[i];

                //var handle = new IntPtr(window.Hwnd);
                var path = workbook.FullName;
                Marshal.ReleaseComObject(workbook);
                yield return new WordWindow(IntPtr.Zero, path);
            }
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(_workbooks);
            Marshal.ReleaseComObject(_windows);
            Marshal.ReleaseComObject(_excel);
        }

    }
}