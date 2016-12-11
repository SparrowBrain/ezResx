using CommandLine;

namespace ezResx.CommandLineData
{
    [Verb("import", HelpText = "Import xlsx into solution, updating or creating resources.")]
    class ImportOptions
    {
        [Option('s', "solution", Required = true, HelpText = "Path to solution file (*.sln)")]
        public string Solution { get; set; }

        [Option('e', "excel", Required = true, HelpText = "Path to excel file (*.xlsx)")]
        public string Excel { get; set; }
    }
}