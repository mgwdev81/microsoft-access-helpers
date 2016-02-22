using System;
using System.IO;

namespace Common
{
    public class FileLogger : ILogger
    {
        string filePath;

        public FileLogger(string filePath)
        {
            this.filePath = filePath;

            var directory = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }

        public void Log(string message)
        {
            using (var writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine(FormatLogMessage(message));
            }
        }

        private string FormatLogMessage(string message)
        {
            return string.Format("{0}|{1}", DateTime.Now.ToString("u"), message);
        }
    }
}
