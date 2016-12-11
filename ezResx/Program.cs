using System;
using CommandLine;
using ezResx.CommandLineData;
using ezResx.Excel;
using ezResx.Resource;
using ezResx.Solution;

namespace ezResx
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Parser.Default.ParseArguments<ExportOptions, ImportOptions, MergeOptions, FullOptions>(args)
                .WithParsed<ExportOptions>(opts => Try(() => { Export(opts.Solution, opts.Excel); }))
                .WithParsed<ImportOptions>(opts => Try(() => { Import(opts.Solution, opts.Excel); }))
                .WithParsed<MergeOptions>(opts => Try(() => { Merge(opts.Solution, opts.Excel); }))
                .WithParsed<FullOptions>(opts => Try(() =>
                {
                    Merge(opts.Solution, opts.Excel);
                    Import(opts.Solution, opts.Excel);
                }));
        }

        private static void Try(Action action)
        {
            try
            {
                action.Invoke();
                Console.WriteLine("Done!");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex);
                Console.ResetColor();
            }
        }
        
        private static void Export(string solutionPath, string translationsXlsx)
        {
            Console.WriteLine("Reading solution resources...");
            var resourceList = new SolutionReader().GetSolutionResources(solutionPath);

            Console.WriteLine("Exporting to xlsx...");
            var excelWriter = new ExcelWriter();
            excelWriter.WriteXlsx(translationsXlsx, resourceList);
        }

        private static void Merge(string solutionPath, string translationsXlsx)
        {
            Console.WriteLine("Reading solution resources...");
            var solutionResources = new SolutionReader().GetSolutionResources(solutionPath);

            Console.WriteLine("Reading xlsx resources...");
            var excelReader = ExcelReader.CreateReader(translationsXlsx);
            var xlsxResources = excelReader.ReadXlsx();

            Console.WriteLine("Merging resources...");
            var resources = new ResourceMerger().MergeResources(solutionResources, xlsxResources);

            Console.WriteLine("Writing resources to xlsx...");
            var excelWriter = new ExcelWriter();
            excelWriter.WriteXlsx(translationsXlsx, resources);
        }

        private static void Import(string solutionPath, string translationsXlsx)
        {
            Console.WriteLine("Reading xlsx resources...");
            var excelReader = ExcelReader.CreateReader(translationsXlsx);
            var resources = excelReader.ReadXlsx();

            Console.WriteLine("Adding resources to solution...");
            new SolutionWriter().AddResourcesToSolution(solutionPath, resources);
        }
    }
}