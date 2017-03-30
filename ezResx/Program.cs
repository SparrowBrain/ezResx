﻿using System;
using CommandLine;
using ezResx.CommandLineData;
using ezResx.Excel;
using ezResx.Resource;
using ezResx.Solution;
using System.IO;
using ezResx.Errors;
using ezResx.Data;

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
            catch (DataLossException except)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Potential data loss");
                Console.WriteLine(except.Message);
                foreach (var lostResource in except.MissingData)
                {
                    Console.WriteLine($"{lostResource.Key.Project} { lostResource.Key.File} {lostResource.Key.Name }");
                }
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
           
        }
        
        private static void Export(string solutionPath, string translationsXlsx)
        {
            Console.WriteLine("Reading solution resources...");
            var resourceList = new SolutionReader().GetSolutionResources(solutionPath);

            Console.WriteLine($"Exporting to xlsx file: {Path.GetFileName(translationsXlsx)}");
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

            Console.WriteLine($"Writing resources to file: {Path.GetFileName(translationsXlsx)}");
            var excelWriter = new ExcelWriter();
            excelWriter.WriteXlsx(translationsXlsx, resources);
        }

        private static void Import(string solutionPath, string translationsXlsx)
        {
            Console.WriteLine("Reading xlsx resources...");
            var excelReader = ExcelReader.CreateReader(translationsXlsx);
            var resources = excelReader.ReadXlsx();

            Console.WriteLine($"Adding resources to solution: {Path.GetFileName(solutionPath)}");                       
            new SolutionWriter().AddResourcesToSolution(solutionPath, resources);
        }
    }
}