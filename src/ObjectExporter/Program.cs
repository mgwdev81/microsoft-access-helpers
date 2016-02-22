using CommandLine;
using Common;
using Microsoft.Office.Interop.Access;
using System;
using System.IO;

namespace ObjectExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            Application application;
            AccessFileProvider accessFileProvider;
            CommandLineOptions options = new CommandLineOptions();
            Logger logger = new Logger();
            
            if (!Parser.Default.ParseArguments(args, options))
            {
                Console.WriteLine(options.GetUsage());
                return;
            }

            if (!Directory.Exists(options.OutputDirectory))
                Directory.CreateDirectory(options.OutputDirectory);
            
            logger.RegisterLogger(new ConsoleLogger());
            logger.RegisterLogger(new FileLogger(Path.Combine(options.OutputDirectory, "ObjectExporter.log")));
                       
            application = new Application();
            accessFileProvider = new AccessFileProvider(options.InputDirectory);

            foreach (FileInfo fileInfo in accessFileProvider.Files)
            {
                application.OpenCurrentDatabase(fileInfo.FullName);

                var objectExporter = new ObjectExporter(application, logger, fileInfo.FullName, options.OutputDirectory);

                logger.Log(string.Format("Beginning export for database: {0}.", fileInfo.FullName));

                if (options.TypesToExport == null)
                {
                    objectExporter.ExportAll();
                }
                else
                {
                    if (options.TypesToExport.Contains("table")) objectExporter.ExportTables();
                    if (options.TypesToExport.Contains("query")) objectExporter.ExportQueries();
                    if (options.TypesToExport.Contains("macro")) objectExporter.ExportMacros();
                    if (options.TypesToExport.Contains("module")) objectExporter.ExportModules();
                }

                logger.Log(string.Format("Completed export for database: {0}.", fileInfo.FullName));

                application.CloseCurrentDatabase();
            }

            application.Quit();
            application = null;
        }
    }
}
