using CfsImportManager.TablesInfo;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CfsImportManager
{
    public class ExcelBase
    {
        public ExcelBase(string excelPath)
        {
            WorkbookDefault = new XLWorkbook(excelPath);
            WorkbookTrimmed = new XLWorkbook(excelPath);
            SettingsWorksheetTrimmed = WorkbookTrimmed.Worksheets.Single(x => x.Name == "Настройки");
            SetCommonTablesCells();
            SetMainTablesCells();
            CreateCommonTablesInfos();
            CreateMainTablesInfos();
        }
        public static XLWorkbook WorkbookDefault { get; set; }
        public static XLWorkbook WorkbookTrimmed { get; set; }
        public static IXLWorksheet SettingsWorksheetTrimmed { get; set; }
        public static List<IXLCell> CommonTablesCells { get; set; } = new List<IXLCell>();
        public static List<IXLCell> MainTablesCells { get; set; } = new List<IXLCell>();
        public static List<CommonTableInfo> CommonTablesInfos { get; set; } = new List<CommonTableInfo>();
        public static List<MainTableInfo> MainTablesInfos { get; set; } = new List<MainTableInfo>();
        private void SetCommonTablesCells()
        {
            var firstCell = SettingsWorksheetTrimmed.Search("Общие таблицы").First().CellBelow();
            var lastCell = firstCell.WorksheetColumn().LastCellUsed();
            CommonTablesCells = SettingsWorksheetTrimmed.Range(firstCell, lastCell).Cells().ToList();
        }
        private void SetMainTablesCells()
        {
            var firstCell = SettingsWorksheetTrimmed.Search("Основная таблица").First().CellRight();
            do
            {
                MainTablesCells.Add(firstCell);
                firstCell = firstCell.CellRight(4);
            }
            while (!firstCell.IsEmpty());
        }
        private void CreateCommonTablesInfos()
        {
            foreach(var tableCell in CommonTablesCells)
            {
                CommonTableInfo tableInfo = new CommonTableInfo();
                tableInfo.Queue = (int)tableCell.CellRight().Value;
                tableInfo.DoublesColumn = tableCell.CellRight(2).Value.ToString();
                tableInfo.IdUpdateColumn = tableCell.CellRight(3).Value.ToString();
                tableInfo.TableName = tableCell.Value.ToString();
                tableInfo.TrimmedWorksheet = GetTrimmedWorksheet(tableCell);
                tableInfo.RowsUsedCount = tableInfo.TrimmedWorksheet.RowsUsed().Count();

                CommonTablesInfos.Add(tableInfo);
            }

        }
        private void CreateMainTablesInfos()
        {
            foreach (var tableCell in MainTablesCells)
            {
                MainTableInfo tableInfo = new MainTableInfo();
                tableInfo.Queue = (int)tableCell.CellRight().Value;
                tableInfo.DoublesColumn = tableCell.CellRight(2).Value.ToString();
                tableInfo.IdUpdateColumn = tableCell.CellRight(3).Value.ToString();
                tableInfo.TableName = tableCell.Value.ToString();
                tableInfo.DefaultWorksheet = WorkbookTrimmed.Worksheets.Single(x => x.Name == tableCell.Value.ToString());
                tableInfo.TrimmedWorksheet = GetTrimmedWorksheet(tableCell);
                tableInfo.RowsUsedCount = tableInfo.TrimmedWorksheet.RowsUsed().Count();
                tableInfo.ReferenceTablesCells = SetRefenceTablesCells((string)tableCell.Value);
                tableInfo.ReferenceTables = SetReferenceTables(tableInfo.ReferenceTablesCells);

                MainTablesInfos.Add(tableInfo);
            }
            List<IXLCell> SetRefenceTablesCells(string mainTableName)
            {
                var firstCell = SettingsWorksheetTrimmed.Search(mainTableName).First().CellBelow();
                var lastCell = firstCell.WorksheetColumn().LastCellUsed();
                return SettingsWorksheetTrimmed.Range(firstCell, lastCell).Cells().ToList();
            }
            List<ReferenceTableInfo> SetReferenceTables(List<IXLCell> tablesCells)
            {
                List<ReferenceTableInfo> referenceTables = new List<ReferenceTableInfo>();
                foreach (var tableCell in tablesCells)
                {
                    ReferenceTableInfo tableInfo = new ReferenceTableInfo();
                    tableInfo.Queue = (int)tableCell.CellRight().Value;
                    tableInfo.TableName = tableCell.Value.ToString();
                    tableInfo.DoublesColumn = tableCell.CellRight(2).Value.ToString();
                    tableInfo.IdUpdateColumn = tableCell.CellRight(3).Value.ToString();
                    tableInfo.DefaultWorksheet = WorkbookTrimmed.Worksheets.Single(x => x.Name == tableCell.Value.ToString());
                    tableInfo.TrimmedWorksheet = GetTrimmedWorksheet(tableCell);
                    tableInfo.RowsUsedCount = tableInfo.TrimmedWorksheet.RowsUsed().Count();

                    referenceTables.Add(tableInfo);
                }
                return referenceTables;
            }
        }
        private IXLWorksheet GetTrimmedWorksheet(IXLCell tableCell)
        {
            //fontName
            //ff92d050 - зеленый
            //ffb2b2b2 - серый
            //Логика усложнена так как цвета глючат.
            var worksheet = WorkbookTrimmed.Worksheets.Single(x => x.Name == tableCell.Value.ToString());
            var firstRowCells = worksheet.Row(1).Cells().ToList();
            foreach (var cell in firstRowCells)
            {
                try
                {
                    if (cell.Style.Fill.BackgroundColor.Color.Name != "ff92d050" && cell.Style.Fill.BackgroundColor.Color.Name != "ffb2b2b2")
                    {
                        cell.WorksheetColumn().Delete();
                    }
                }
                catch
                {
                    cell.WorksheetColumn().Delete();
                    continue;
                }
            }
            /*var a = worksheet.Row(1).Cells();*/
            return worksheet;
        }
    }
}
