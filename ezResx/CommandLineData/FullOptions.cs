using CommandLine;

namespace ezResx.CommandLineData
{
    [Verb("full",
        HelpText =
            "Full update procedure. Updates input excel by merging with solution. Updates/creates/deletes solution resources."
        )]
    internal class FullOptions
    {
        [Option('s', "solution", Required = true, HelpText = "Path to solution file (*.sln)")]
        public string Solution { get; set; }

        [Option('e', "excel", Required = true, HelpText = "Xlsx file containing latest translations.")]
        public string Excel { get; set; }
    }
}