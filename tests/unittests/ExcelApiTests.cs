using System.Linq;
using System.Threading;
using FluentAssertions;
using OfficeInteropLib;
using OfficeInteropLib.Common;
using Xunit;
using Xunit.Abstractions;

namespace UnitTests
{
    public class ExcelApiTests
    {
        private readonly ITestOutputHelper _logger;

        public ExcelApiTests(ITestOutputHelper logger)
        {
            _logger = logger;
        }

        [Fact]
        public void Version_ShouldReturnVersion()
        {
            using var sut = new ExcelApi();

            var version = sut.Version;

            version.Should().NotBeEmpty();
            
            _logger.WriteLine(version);
        }

        [Fact]
        public void GetWindows_ShouldNotBeEmpty_WhenThereIsExcelDocumentOpened()
        {
            //var excelWorkbook = Application.Launch("test.xlsx").WaitWhileMainHandleIsMissing();
            Thread.Sleep(400);

            using var sut = new ExcelApi();

            var windows = sut.GetWindows();

            windows.Should().NotBeEmpty();

            windows.ToList().ForEach(x=> _logger.WriteLine($"{x.Handle} : {x.DocumentPath}"));
            //excelWorkbook.Kill();
        }
    }
}