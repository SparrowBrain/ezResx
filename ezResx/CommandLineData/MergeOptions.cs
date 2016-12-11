using CommandLine;

namespace ezResx.CommandLineData
{
    [Verb("merge", HelpText = "Merge information from solution with modified xlsx file, updating the xlsx.")]
    class MergeOptions
    {
        [Option('s', "solution", Required = true, HelpText = "Path to solution file (*.sln)")]
        public string Solution { get; set; }

        [Option('e', "excel", Required = true, HelpText = "Xlsx file containing latest translations. Will be overwritten!")]
        public string Excel { get; set; }
    }
}