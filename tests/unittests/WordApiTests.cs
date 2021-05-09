using System;
using System.Linq;
using FluentAssertions;
using OfficeInteropLib;
using OfficeInteropLib.Common;
using Xunit;
using Xunit.Abstractions;

using static UnitTests.Testing;

namespace UnitTests
{
    public class WordApiTests
    {
        private readonly ITestOutputHelper _logger;

        public WordApiTests(ITestOutputHelper logger)
        {
            _logger = logger;
        }
        
        [Fact]
        public void Version_ShouldReturnVersion()
        {
            using var sut = new WordApi();

            var version = sut.Version;

            version.Should().NotBeEmpty();
            
            _logger.WriteLine(version);
        }

        [Fact]
        public void GetWindows_ShouldNotBeEmptyWhenThereIsAWordDocumentAvailable()
        {
            Application.Launch("test.docx").WaitWhileMainHandleIsMissing();
            Retry.WhileFalse(WordApi.HasRunningInstances, 10, 500);

            using var sut = new WordApi();

            var windows = sut.GetWindows();

            windows.Should().NotBeEmpty();
            var window = windows.FirstOrDefault();
            window.DocumentPath.Should().NotBeEmpty();
            window.Handle.Should().NotBe(IntPtr.Zero);

            windows.ToList().ForEach(x=> _logger.WriteLine($"{x.Handle} : {x.DocumentPath}"));

            KillProcesses("Excel");
        }


        [Fact]
        public void GetWindows2_ShouldNotBeEmptyWhenThereIsAWordDocumentAvailable()
        {
            Application.Launch("test.docx").WaitWhileMainHandleIsMissing();
            Retry.WhileFalse(WordApi.HasRunningInstances, 10, 500);

            var windows = WordApi.GetWindows2();

            windows.Should().NotBeEmpty();
            var window = windows.FirstOrDefault();
            window.DocumentPath.Should().NotBeEmpty();
            window.Handle.Should().NotBe(IntPtr.Zero);

            windows.ToList().ForEach(x=> _logger.WriteLine($"{x.Handle} : {x.DocumentPath}"));

            KillProcesses("Excel");
        }
    }
}
