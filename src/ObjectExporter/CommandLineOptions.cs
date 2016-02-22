using CommandLine;
using CommandLine.Text;
using System.Collections.Generic;

namespace ObjectExporter
{
    class CommandLineOptions
    {
        [Option('i', "inputDirectory", Required = true, 
            HelpText = "Directory containing MS Access database(s) to search.")]
        public string InputDirectory { get; set; }

        [Option('o', "outputDirectory", Required = true, 
            HelpText = "Directory to save exported database object files. Directory will be created if it doesn't exist.")]
        public string OutputDirectory { get; set; }

        [OptionList('e', "typesToExport",
            HelpText = "Object types to export. Can be one or more of: table, query, macro, module.")]
        public IList<string> TypesToExport { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
