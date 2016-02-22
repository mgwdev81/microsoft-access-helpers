using System.Collections.Generic;

namespace Common
{
    public class Logger
    {
        private List<ILogger> loggers;

        public Logger()
        {
            loggers = new List<ILogger>();
        }

        public void RegisterLogger(ILogger logger)
        {
            loggers.Add(logger);
        }

        public void Log(string message)
        {
            loggers.ForEach(l => l.Log(message));
        }
    }
}
