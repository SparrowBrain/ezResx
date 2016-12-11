using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ezResx.Data;

namespace ezResx.Excel
{
    internal class ExcelWriter : ExcelService
    {
        private XLWorkbook _workBook;
        private IXLWorksheet _sheet;
        private IXLColumn _projectColumn;
        private IXLColumn _fileColumn;
        private IXLColumn _nameColumn;
        private IXLColumn _defaultCultureColumn;

        public ExcelWriter()
        {
            CreateSheetAndHeaders();
        }

        public void WriteXlsx(string filepath, List<ResourceItem> resourceList)
        {
            foreach (var resource in resourceList)
            {
                var currentRow = _sheet.LastRowUsed().RowBelow();

                WriteKeyValues(currentRow, resource);
                WriteDefaultValue(currentRow, resource);
                WriteLocaleValues(currentRow, resource);
            }

            _sheet.FirstRow().CellsUsed().Style.Fill.BackgroundColor = XLColor.Yellow;
            _workBook.SaveAs(filepath);
        }
        
        private void WriteKeyValues(IXLRow currentRow, ResourceItem resource)
        {
            currentRow.Cell(_projectColumn.ColumnNumber()).Value = resource.Key.Project;
            currentRow.Cell(_fileColumn.ColumnNumber()).Value = resource.Key.File;
            currentRow.Cell(_nameColumn.ColumnNumber()).Value = resource.Key.Name;
        }

        private void WriteDefaultValue(IXLRow currentRow, ResourceItem resource)
        {
            string defaultValue;
            if (!resource.Values.TryGetValue(DefaultCultureColumn, out defaultValue))
            {
                throw new Exception($"Resource default culture not found for {resource.Key.Project} {resource.Key.File} {resource.Key.Name}");
            }

            currentRow.Cell(_defaultCultureColumn.ColumnNumber()).Value = defaultValue;
        }

        private void WriteLocaleValues(IXLRow currentRow, ResourceItem resource)
        {
            foreach (var value in resource.Values)
            {
                if (value.Key == DefaultCultureColumn)
                {
                    continue;
                }

                var columnHeader = _sheet.FirstRow().Cells().FirstOrDefault(x => x.Value.ToString() == value.Key);
                if (columnHeader == null)
                {
                    _sheet.FirstRow().LastCellUsed().CellRight().Value = value.Key;
                    columnHeader = _sheet.LastColumnUsed().FirstCell();
                }

                currentRow.Cell(columnHeader.WorksheetColumn().ColumnNumber()).Value = value.Value;
            }
        }
        
        private void CreateSheetAndHeaders()
        {
            _workBook = new XLWorkbook();
            _sheet = _workBook.Worksheets.Add(ResourceSheetName);
            _projectColumn = CreateColumn(_sheet, ProjectColumnName);
            _fileColumn = CreateColumn(FileColumnName, _projectColumn);
            _nameColumn = CreateColumn(NameColumnName, _fileColumn);
            _defaultCultureColumn = CreateColumn(DefaultCultureColumn, _nameColumn);
            CreateColumn("da", _defaultCultureColumn);
        }

        private IXLColumn CreateColumn(IXLWorksheet sheet, string name)
        {
            var projectColumn = sheet.FirstColumn();
            projectColumn.FirstCell().Value = name;
            return projectColumn;
        }

        private IXLColumn CreateColumn(string name, IXLColumn previousColumn)
        {
            var newColumn = previousColumn.ColumnRight();
            newColumn.FirstCell().Value = name;
            return newColumn;
        }
    }
}