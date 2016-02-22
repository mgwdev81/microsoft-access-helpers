using CommandLine;
using CommandLine.Text;

namespace ApplicationOption
{
    class CommandLineOptions
    {
        [Option('i', "inputDirectory", Required = true, 
            HelpText = "Directory containing MS Access database(s) to search.")]
        public string InputDirectory { get; set; }

        [Option('o', "outputDirectory", Required = true,
            HelpText = "Directory to save log file to. Directory will be created if it doesn't exist.")]
        public string OutputDirectory { get; set; }

        [Option('p', "optionName", Required = true, 
            HelpText = "Name of the option to set.")]
        public string AccessOption { get; set; }

        [Option('v', "optionValue", Required = true,
            HelpText = "Value to set.")]
        public string AccessOptionValue { get; set; }

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
