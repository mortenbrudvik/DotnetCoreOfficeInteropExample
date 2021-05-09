using System.Diagnostics;
using System.Linq;
using System.Threading;
using FluentAssertions;
using OfficeInteropLib;
using OfficeInteropLib.Common;
using Xunit;
using Xunit.Abstractions;

using static UnitTests.Testing;

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
            Application.Launch("test.xlsx").WaitWhileMainHandleIsMissing();
            Retry.WhileFalse(ExcelApi.HasRunningInstances, 10, 500);

            using var sut = new ExcelApi();

            var windows = sut.GetWindows();

            windows.Should().NotBeEmpty();

            windows.ToList().ForEach(x=> _logger.WriteLine(x.ToString()));

            KillProcesses("EXCEL");
        }

        [Fact]
        public void GetWindows2_ShouldNotBeEmpty_WhenThereIsExcelDocumentOpened()
        {
            Application.Launch("test.xlsx");
            Retry.WhileFalse(ExcelApi.HasRunningInstances, 10, 500);

            var windows = ExcelApi.GetWindows2();

            windows.Should().NotBeEmpty();

            windows.ToList().ForEach(x=> _logger.WriteLine(x.ToString()));
            KillProcesses("EXCEL");
        }
    }
}