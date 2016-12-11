using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Resources;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;
using ezResx.Data;
using ezResx.Excel;
using ezResx.Resource;
using ezResx.Solution;
using Microsoft.Build.Evaluation;


namespace ezResx
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var solutionPath = @"C:\Users\Qwx\Source\Repos\MachoKey\MachoKey.sln";
            var translationsXlsx = "Translations.xlsx";

            var resourceList = new SolutionReader().GetSolutionResources(solutionPath);

            // Export excel
            var excelWriter = new ExcelWriter();
            excelWriter.WriteXlsx(translationsXlsx, resourceList);

            // Read xlsx
            var excelReader = ExcelReader.CreateReader(translationsXlsx);
            var resources = excelReader.ReadXlsx();

            // merge
            var solutionResources = resourceList;
            var xlsxResources = resources;
            
            resources = new ResourceMerger().MergeResources(solutionResources, xlsxResources);


            // Import xlsx
            new SolutionWriter().AddResourcesToSolution(solutionPath, resources);


            Console.ReadKey();
        }
    }
}