using System.Linq;
using FluentAssertions;
using OfficeInteropLib;
using OfficeInteropLib.Common;
using Xunit;
using Xunit.Abstractions;

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
            var wordDoc = Application.Launch("test.docx").WaitWhileMainHandleIsMissing();

            using var sut = new WordApi();

            var windows = sut.GetWindows();

            windows.Should().NotBeEmpty();

            windows.ToList().ForEach(x=> _logger.WriteLine($"{x.Handle} : {x.DocumentPath}"));

            wordDoc.Kill();
        }
    }
}
