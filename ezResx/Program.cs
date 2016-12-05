using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;
using Microsoft.Build.Evaluation;


namespace ezResx
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //\u0022([^\u0022]*.csproj)\u0022
            //var solutionRegex = new Regex(@"Project.*\""([^""]+\.csproj)\"",.*EndProject", RegexOptions.Compiled | RegexOptions.CultureInvariant);

            var solutionPath = @"C:\Users\Qwx\Source\Repos\MachoKey\MachoKey.sln";
            var solutionDirectory = Path.GetDirectoryName(solutionPath);
            var solutionRegex = new Regex(@"\u0022([^\u0022]+\.csproj)\u0022");//, RegexOptions.Compiled | RegexOptions.CultureInvariant);
            string solution;
            using (var reader = new StreamReader(solutionPath))
            {
                solution = reader.ReadToEnd();
                
            }

            var projects = new List<string>();
            var matches = solutionRegex.Matches(solution);
            foreach (Match match in matches)
            {
                var projectPath = match.Groups[1].Value;
                projects.Add(projectPath);
            }

            var resourceList = new List<ResourceItem>();
            foreach (var projectPath in projects)
            {
                var project = new Project(Path.Combine(solutionDirectory, projectPath));
                        //@"C:\src\tfs\P2P Mobile Wallet\MobilePayLoyalty\DB.Mobeco.MobilePay.Loyalty.Infrastructure\DB.Mobeco.MobilePay.Loyalty.Infrastructure.csproj");
                        
                var items = project.Items.Where(x => x.EvaluatedInclude.EndsWith(".resx"));

                foreach (var item in items)
                {
                    var filePath = Path.Combine(project.DirectoryPath, item.EvaluatedInclude);
                    var withoutExtension = Path.GetFileNameWithoutExtension(item.EvaluatedInclude);
                    var locale = withoutExtension.Contains('.') ? withoutExtension.Split('.').Last() : "default-culture";
                    if (locale != "default-culture")
                    {
                        continue;
                    }

                    var file = withoutExtension.Split('.').First();

                    foreach (var dataElement in XElement.Load(filePath).Elements("data"))
                    {
                        var resourceKey = new ResourceKey
                        {
                            Project = projectPath,
                            // TODO file should retain path within project
                            File = file,
                            Name = dataElement.Attribute("name").Value
                        };

                        var resourceItem = new ResourceItem
                        {
                            Key = resourceKey,
                            Values = new Dictionary<string, string>()
                        };

                        resourceItem.Values[locale] = dataElement.Element("value").Value;

                        resourceList.Add(resourceItem);

                    }

                }


                foreach (var item in items)
                {
                    var filePath = Path.Combine(project.DirectoryPath, item.EvaluatedInclude);
                    var withoutExtension = Path.GetFileNameWithoutExtension(item.EvaluatedInclude);
                    var locale = withoutExtension.Contains('.') ? withoutExtension.Split('.').Last() : "default-culture";

                    if (locale == "default-culture")
                    {
                        continue;
                    }

                    var file = withoutExtension.Split('.').First();

                    foreach (var dataElement in XElement.Load(filePath).Elements("data"))
                    {
                        var resourceKey = new ResourceKey
                        {
                            Project = projectPath,
                            // TODO file should retain path within project
                            File = file,
                            Name = dataElement.Attribute("name").Value
                        };

                        var resourceItem = resourceList.FirstOrDefault(x => Equals(x.Key, resourceKey));
                        if (resourceItem == null)
                        {
                            Console.WriteLine(
                                $"No invariant culture resource found for {resourceKey.File} {resourceKey.Name}");
                            continue;
                        }

                        resourceItem.Values[locale] = dataElement.Element("value").Value;

                        Console.WriteLine("{0} = {1}", dataElement.Attribute("name").Value,
                            dataElement.Element("value").Value);
                    }
                }
            }

            // Export excel

            ExportXlsx(resourceList);

            // Read xlsx

            var resources = ReadXlsx("Translations.xlsx");

            // merge
            var solutionResources = resourceList;
            var xlsxResources = resources;

            MergeResources(solutionResources, xlsxResources);


            Console.ReadKey();
        }

        private static void ExportXlsx(List<ResourceItem> resourceList)
        {
            var workBook = new XLWorkbook();
            var sheet = workBook.Worksheets.Add("Resources");
            var projectColumn = CreateColumn(sheet, "Project");
            var fileColumn = CreateColumn(projectColumn, "File");
            var nameColumn = CreateColumn(fileColumn, "Name");
            var defaultCulture = CreateColumn(nameColumn, "default-culture");
            var daColumn = CreateColumn(defaultCulture, "da");
            sheet.FirstRow().Style.Fill.BackgroundColor = XLColor.Yellow;

            foreach (var resource in resourceList)
            {
                sheet.LastRowUsed().RowBelow().Cell(projectColumn.ColumnNumber()).Value = resource.Key.Project;
                sheet.LastRowUsed().Cell(fileColumn.ColumnNumber()).Value = resource.Key.File;
                sheet.LastRowUsed().Cell(nameColumn.ColumnNumber()).Value = resource.Key.Name;
                string defaultValue;
                if (!resource.Values.TryGetValue("default-culture", out defaultValue))
                {
                    throw new Exception(
                        $"Resource default culture not found for {resource.Key.File} {resource.Key.Name}");
                }

                sheet.LastRowUsed().Cell(defaultCulture.ColumnNumber()).Value = defaultValue;

                foreach (var value in resource.Values)
                {
                    if (value.Key == "default-culture")
                    {
                        continue;
                    }

                    var columnHeader = sheet.FirstRow().Cells().FirstOrDefault(x => x.Value.ToString() == value.Key);
                    if (columnHeader == null)
                    {
                        sheet.FirstRow().LastCellUsed().CellRight().Value = value.Key;
                        columnHeader = sheet.LastColumnUsed().FirstCell();
                    }

                    sheet.LastRowUsed().Cell(columnHeader.WorksheetColumn().ColumnNumber()).Value = value.Value;
                }
            }

            workBook.SaveAs("Translations.xlsx");
        }

        private static List<ResourceItem> ReadXlsx(string file)
        {
            var workBook = new XLWorkbook(file);
            var sheet = workBook.Worksheet("Resources");

            var projectColumn = sheet.FirstRow().Cells().FirstOrDefault(x => x.Value.ToString() == "Project").WorksheetColumn();
            var fileColumn = sheet.FirstRow().Cells().FirstOrDefault(x => x.Value.ToString() == "File").WorksheetColumn();
            var nameColumn = sheet.FirstRow().Cells().FirstOrDefault(x => x.Value.ToString() == "Name").WorksheetColumn();

            var localeHeaders = new List<IXLCell>();
            var defaultHeader = sheet.FirstRow().Cells().First(x => x.Value.ToString() == "default-culture");
            localeHeaders.Add(defaultHeader);
            var rightCell = defaultHeader.CellRight();
            string locale;
            while (rightCell.TryGetValue<string>(out locale) && !string.IsNullOrWhiteSpace(locale))
            {
                localeHeaders.Add(rightCell);
                rightCell = rightCell.CellRight();
            }

            var resources = new List<ResourceItem>();
            var firstRow = true;
            foreach (var row in sheet.RowsUsed())
            {
                if (firstRow)
                {
                    firstRow = false;
                    continue;
                }

                var project = row.Cell(projectColumn.ColumnNumber()).Value.ToString();
                var file1 = row.Cell(fileColumn.ColumnNumber()).Value.ToString();
                var name = row.Cell(nameColumn.ColumnNumber()).Value.ToString();

                var resource = new ResourceItem()
                {
                    Key = new ResourceKey {Project = project, File = file1, Name = name},
                    Values = new Dictionary<string, string>()
                };


                foreach (var localeHeader in localeHeaders)
                {
                    string value;
                    if (row.Cell(localeHeader.WorksheetColumn().ColumnNumber()).TryGetValue(out value) &&
                        (!string.IsNullOrWhiteSpace(value) || localeHeader.Value.ToString() == "default-culture"))
                    {
                        resource.Values[localeHeader.Value.ToString()] = value;
                    }
                }

                resources.Add(resource);
            }
            return resources;
        }

        private static IEnumerable<ResourceItem> MergeResources(List<ResourceItem> solutionResources, List<ResourceItem> xlsxResources)
        {
            foreach (var solutionResource in solutionResources)
            {
                var xlsxResource = xlsxResources.FirstOrDefault(x => x.Key.Equals(solutionResource.Key));
                if (xlsxResource != null)
                {
                    foreach (var value in xlsxResource.Values)
                    {
                        solutionResource.Values[value.Key] = value.Value;
                    }

                    xlsxResources.Remove(xlsxResource);
                }
            }

            if (xlsxResources.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Translations will be lost:");
                foreach (var lostResource in xlsxResources)
                {
                    Console.WriteLine($"{lostResource.Key.File} {lostResource.Key.Name}");
                }

                Console.ResetColor();
            }

            return solutionResources;
        }

        private static IXLColumn CreateColumn(IXLWorksheet sheet, string name)
        {
            var projectColumn = sheet.FirstColumn();
            projectColumn.FirstCell().Value = name;
            return projectColumn;
        }

        private static IXLColumn CreateColumn(IXLColumn previousColumn, string name)
        {
            var newColumn = previousColumn.ColumnRight();
            newColumn.FirstCell().Value = name;
            return newColumn;
        }
    }

    internal class ResourceKey
    {
        public string Project { get; set; }

        public string File { get; set; }

        public string Name { get; set; }

        public override bool Equals(object obj)
        {
            var other = obj as ResourceKey;
            if (other == null)
            {
                return false;
            }

            // TODO dictionary and hash
            return Project == other.Project && File == other.File && Name == other.Name;
        }
    }
    

    internal class ResourceItem
    {
        public ResourceKey Key { get; set; }

        // TODO dictionary
        public IDictionary<string, string> Values { get; set; }
    }
}