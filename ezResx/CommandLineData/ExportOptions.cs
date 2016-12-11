using CommandLine;

namespace ezResx.CommandLineData
{
    [Verb("export", HelpText = "Export solution resources into xlsx file.")]
    class ExportOptions
    {
        [Option('s', "solution", Required = true, HelpText = "Path to solution file (*.sln)")]
        public string Solution { get; set; }

        [Option('e', "excel", Default = "translations.xlsx", HelpText = "Path to excel file (*.xlsx)")]
        public string Excel { get; set; }
    }
}