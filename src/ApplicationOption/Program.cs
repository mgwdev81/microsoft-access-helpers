using CommandLine;
using Common;
using Microsoft.Office.Interop.Access;
using System;
using System.IO;

namespace ApplicationOption
{
    class Program
    {
        static void Main(string[] args)
        {
            Application application;
            AccessFileProvider accessFileProvider;
            CommandLineOptions options = new CommandLineOptions();
            Logger logger = new Logger(); ;

            if (!Parser.Default.ParseArguments(args, options))
            {
                Console.WriteLine(options.GetUsage());
                return;
            }

            if (!Directory.Exists(options.OutputDirectory))
                Directory.CreateDirectory(options.OutputDirectory);

            logger.RegisterLogger(new ConsoleLogger());
            logger.RegisterLogger(new FileLogger(Path.Combine(options.OutputDirectory, "SetOption.log")));

            application = new Application();
            accessFileProvider = new AccessFileProvider(options.InputDirectory);

            foreach (FileInfo fileInfo in accessFileProvider.Files)
            {
                application.OpenCurrentDatabase(fileInfo.FullName);
                
                var optionValueBefore = ApplicationOption.GetOption(application, options.AccessOption);
                ApplicationOption.SetOption(application, options.AccessOption, options.AccessOptionValue);
                var optionValueAfter = ApplicationOption.GetOption(application, options.AccessOption);

                application.CloseCurrentDatabase();

                logger.Log(string.Format("{0}: Option '{1}' changed from '{2}' to '{3}'.", 
                    fileInfo.FullName, options.AccessOption, optionValueBefore, optionValueAfter));
            }
        }
    }
}
