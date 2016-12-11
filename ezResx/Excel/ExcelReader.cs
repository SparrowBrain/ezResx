using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ezResx.Data;

namespace ezResx.Excel
{
    internal class ExcelReader : ExcelService
    {
        private ExcelReader()
        {
        }

        private IXLWorksheet Sheet { get; set; }

        private IXLColumn ProjectColumn { get; set; }

        private IXLColumn FileColumn { get; set; }

        private IXLColumn NameColumn { get; set; }

        public static ExcelReader CreateReader(string filePath)
        {
            var workBook = new XLWorkbook(filePath);
            var sheet = workBook.Worksheet(ResourceSheetName);

            var projectColumn = GetColumn(ProjectColumnName, sheet);
            var fileColumn = GetColumn(FileColumnName, sheet);
            var nameColumn = GetColumn(NameColumnName, sheet);

            var excelReader = new ExcelReader
            {
                Sheet = sheet,
                ProjectColumn = projectColumn,
                FileColumn = fileColumn,
                NameColumn = nameColumn
            };

            return excelReader;
        }

        public List<ResourceItem> ReadXlsx()
        {
            var localeHeaders = GetLocaleHeaders();

            var resources = new List<ResourceItem>();
            var headersRow = true;
            foreach (var row in Sheet.RowsUsed())
            {
                if (headersRow)
                {
                    headersRow = false;
                    continue;
                }

                var resource = CreateResourceItem(row);
                FillLocaleValues(resource, row, localeHeaders);

                resources.Add(resource);
            }

            return resources;
        }

        private List<IXLCell> GetLocaleHeaders()
        {
            var localeHeaders = new List<IXLCell>();
            var defaultHeader = Sheet.FirstRow().Cells().FirstOrDefault(x => x.Value.ToString() == DefaultCultureColumn);
            if (defaultHeader == null)
            {
                throw new Exception("No default culture header found");
            }

            localeHeaders.Add(defaultHeader);
            var rightCell = defaultHeader.CellRight();
            string locale;
            while (rightCell.TryGetValue(out locale) && !string.IsNullOrWhiteSpace(locale))
            {
                localeHeaders.Add(rightCell);
                rightCell = rightCell.CellRight();
            }
            return localeHeaders;
        }

        private ResourceItem CreateResourceItem(IXLRow row)
        {
            var project = row.Cell(ProjectColumn.ColumnNumber()).Value.ToString();
            var file = row.Cell(FileColumn.ColumnNumber()).Value.ToString();
            var name = row.Cell(NameColumn.ColumnNumber()).Value.ToString();

            var resource = new ResourceItem
            {
                Key = new ResourceKey {Project = project, File = file, Name = name},
                Values = new Dictionary<string, string>()
            };
            return resource;
        }

        private static void FillLocaleValues(ResourceItem resource, IXLRow row, List<IXLCell> localeHeaders)
        {
            foreach (var localeHeader in localeHeaders)
            {
                string value;
                if (row.Cell(localeHeader.WorksheetColumn().ColumnNumber()).TryGetValue(out value) &&
                    (!string.IsNullOrWhiteSpace(value) || localeHeader.Value.ToString() == DefaultCultureColumn))
                {
                    resource.Values[localeHeader.Value.ToString()] = value;
                }
            }
        }

        private static IXLColumn GetColumn(string columnName, IXLWorksheet sheet)
        {
            var projectColumn =
                sheet.FirstRow().Cells().FirstOrDefault(x => x.Value.ToString() == columnName)?.WorksheetColumn();
            if (projectColumn == null)
            {
                throw new Exception(columnName + " column not found");
            }
            return projectColumn;
        }
    }
}