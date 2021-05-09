using System;
using LanguageExt;
using OfficeInteropLib.Common;
using OfficeInteropLib.Extensions;
using Xunit.Abstractions;

namespace UnitTests.Extensions
{
    public class TestLoggerExt
    {
        public static class LoggerServiceExt
        {
            public static Unit TimeInSeconds<T>(ITestOutputHelper logger, Action operation) =>
                TimeInSeconds(logger, operation.ToFunc());

            public static T TimeInSeconds<T>(ITestOutputHelper logger, Func<T> operation) =>
                Measure.TimeInSeconds(operation).Map((timeInSeconds, result) =>
                {
                    logger.WriteLine($"Execution time: {timeInSeconds} seconds");
                    return result;
                });
        }
    }
}