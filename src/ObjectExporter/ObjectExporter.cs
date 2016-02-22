using Common;
using Microsoft.Office.Interop.Access;
using Dao = Microsoft.Office.Interop.Access.Dao;
using System;
using System.IO;

namespace ObjectExporter
{
    class ObjectExporter
    {
        public Application application { get; private set; }
        private Logger logger;

        private string outputDirectory;
        private string databaseFileName;
        
        public ObjectExporter(Application application, Logger logger, string databasePath, string outputDirectory)
        {
            this.outputDirectory = outputDirectory;
            this.logger = logger;
            this.application = application;

            databaseFileName = Path.GetFileName(databasePath);
        }

        public void ExportAll()
        {
            ExportMacros();
            ExportModules();
            ExportQueries();
            ExportTables();
        }

        // TODO: Determine what type of table we are dealing with e.g. normal local table, linked table and export appropriately.
        public int ExportTables()
        {
            int tablesExported = 0;
            string textToWrite;
            string fileName;
            string filePath;

            Dao.Database database = application.CurrentDb();

            foreach (Dao.TableDef tableDef in database.TableDefs)
            {
                // Don't export system tables.
                if (tableDef.Name.StartsWith("MSys")) continue;

                fileName = string.Format("{0}.table.{1}.txt", databaseFileName, tableDef.Name);
                filePath = Path.Combine(outputDirectory, fileName);

                textToWrite = "";
                textToWrite += string.Format("Name:            {0}", tableDef.Name) + Environment.NewLine;
                textToWrite += string.Format("Connect:         {0}", tableDef.Connect) + Environment.NewLine;
                textToWrite += string.Format("SourceTableName: {0}", tableDef.SourceTableName);

                WriteToFile(filePath, textToWrite);
                tablesExported++;
            }

            database = null;

            logger.Log(string.Format("Exported {0} tables.", tablesExported));

            return tablesExported;
        }

        public int ExportQueries()
        {
            int queriesExported = 0;
            string textToWrite;
            string fileName;
            string filePath;

            Dao.Database database = application.CurrentDb();

            foreach (Dao.QueryDef queryDef in database.QueryDefs)
            {
                fileName = string.Format("{0}.query.{1}.txt", databaseFileName, queryDef.Name);
                filePath = Path.Combine(outputDirectory, fileName);

                textToWrite = "";
                textToWrite += string.Format("Name:    {0}", queryDef.Name) + Environment.NewLine;
                textToWrite += string.Format("Connect: {0}", queryDef.Connect) + Environment.NewLine;
                textToWrite += string.Format("SQL:     {0}", queryDef.SQL);

                WriteToFile(filePath, textToWrite);
                queriesExported++;
            }

            database = null;

            logger.Log(string.Format("Exported {0} queries.", queriesExported));

            return queriesExported;
        }

        public int ExportMacros()
        {
            int macrosExported = 0;
            string fileName;
            string filePath;

            var currentProject = application.CurrentProject;

            foreach(AccessObject macro in currentProject.AllMacros)
            {
                fileName = string.Format("{0}.macro.{1}.txt", databaseFileName, macro.Name);
                filePath = Path.Combine(outputDirectory, fileName);
                application.SaveAsText(AcObjectType.acMacro, macro.Name, filePath);
                macrosExported++;
            }

            currentProject = null;

            logger.Log(string.Format("Exported {0} macros.", macrosExported));

            return macrosExported;
        }

        public int ExportModules()
        {
            int modulesExported = 0;
            string fileName;
            string filePath;

            var currentProject = application.CurrentProject;

            foreach (AccessObject module in currentProject.AllModules)
            {
                fileName = string.Format("{0}.module.{1}.txt", databaseFileName, module.Name);
                filePath = Path.Combine(outputDirectory, fileName);
                application.SaveAsText(AcObjectType.acModule, module.Name, filePath);
                modulesExported++;
            }

            currentProject = null;

            logger.Log(string.Format("Exported {0} modules.", modulesExported));

            return modulesExported; 
        }

        private void WriteToFile(string filePath, string text)
        {
            using (var writer = new StreamWriter(filePath, false))
            {
                writer.Write(text);
            }
        }
    }
}
