using System.Diagnostics;
using System.Threading;

namespace UnitTests
{
    public static class Testing
    {
        public static void KillProcesses(string name)
        {
            foreach (var process in Process.GetProcessesByName(name))
            {
                process.Kill();
            }

            Thread.Sleep(100);
        }
    }
}