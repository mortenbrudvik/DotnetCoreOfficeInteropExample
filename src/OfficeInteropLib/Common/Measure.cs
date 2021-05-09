using System;
using System.Diagnostics;
using LanguageExt;
using OfficeInteropLib.Extensions;

namespace OfficeInteropLib.Common
{
    public static class Measure
    {
        public static (decimal seconds, Unit result) TimeInSeconds(Action operation) =>
            TimeInSeconds(operation.ToFunc());

        public static (decimal seconds, T result) TimeInSeconds<T>(Func<T> operation)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            
            var result = operation();

            stopwatch.Stop();
            var time =  Math.Round((decimal)stopwatch.ElapsedMilliseconds/1000, 3);

            return (time, result);
        }
    }
}