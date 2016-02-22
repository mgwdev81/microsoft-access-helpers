using System;

namespace Common
{
    public class ConsoleLogger : ILogger
    {
        public void Log(string message)
        {
            Console.WriteLine(FormatLogMessage(message));
        }

        private string FormatLogMessage(string message)
        {
            return string.Format("[{0}] {1}", DateTime.Now.ToString("u"), message);
        }
    }
}
