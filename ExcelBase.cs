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
            Workbook = new XLWorkbook(excelPath);
            /*Workbook.RecalculateAllFormulas();*/

            SettingsWorksheet = Workbook.Worksheets.Single(x => x.Name == "Настройки");
            SetCommonTablesCells();
            SetMainTablesCells();
            CreateCommonTablesInfos();
            CreateMainTablesInfos();
        }
        public static XLWorkbook Workbook { get; set; }
        public static IXLWorksheet SettingsWorksheet { get; set; }
        public static List<IXLCell> CommonTablesCells { get; set; } = new List<IXLCell>();
        public static List<IXLCell> MainTablesCells { get; set; } = new List<IXLCell>();
        public static List<CommonTableInfo> CommonTablesInfos { get; set; } = new List<CommonTableInfo>();
        public static List<MainTableInfo> MainTablesInfos { get; set; } = new List<MainTableInfo>();
        private void SetCommonTablesCells()
        {
            var firstCell = SettingsWorksheet.Search("Общие таблицы").First().CellBelow();
            var lastCell = firstCell.WorksheetColumn().LastCellUsed();
            CommonTablesCells = SettingsWorksheet.Range(firstCell, lastCell).Cells().ToList();
        }
        private void SetMainTablesCells()
        {
            var firstCell = SettingsWorksheet.Search("Основная таблица").First().CellRight();
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
                tableInfo.TableCell = tableCell;
                tableInfo.Queue = (int)tableCell.CellRight().Value;
                tableInfo.Type = TableInfoBase.TableType.Common;
                tableInfo.DoublesColumn = tableCell.CellRight(2).Value.ToString();
                tableInfo.IdUpdateColumn = tableCell.CellRight(3).Value.ToString();
                tableInfo.TableName = tableCell.Value.ToString();
                tableInfo.Worksheet = Workbook.Worksheets.Single(x => x.Name == tableCell.Value.ToString());
                tableInfo.RowsUsedCount = tableInfo.Worksheet.RowsUsed().Count();

                CommonTablesInfos.Add(tableInfo);
            }
        }
        private void CreateMainTablesInfos()
        {
            foreach (var tableCell in MainTablesCells)
            {
                MainTableInfo tableInfo = new MainTableInfo();
                tableInfo.TableCell = tableCell;
                tableInfo.Queue = (int)tableCell.CellRight().Value;
                tableInfo.Type = TableInfoBase.TableType.Main;
                tableInfo.DoublesColumn = tableCell.CellRight(2).Value.ToString();
                tableInfo.IdUpdateColumn = tableCell.CellRight(3).Value.ToString();
                tableInfo.TableName = tableCell.Value.ToString();
                tableInfo.Worksheet = Workbook.Worksheets.Single(x => x.Name == tableCell.Value.ToString());
                tableInfo.RowsUsedCount = tableInfo.Worksheet.RowsUsed().Count();
                tableInfo.ReferenceTablesCells = SetRefenceTablesCells((string)tableCell.Value);
                tableInfo.ReferenceTables = SetReferenceTables(tableInfo.ReferenceTablesCells);

                MainTablesInfos.Add(tableInfo);
            }
            List<IXLCell> SetRefenceTablesCells(string mainTableName)
            {
                var firstCell = SettingsWorksheet.Search(mainTableName).First().CellBelow();
                var lastCell = firstCell.WorksheetColumn().LastCellUsed();
                return SettingsWorksheet.Range(firstCell, lastCell).Cells().ToList();
            }
            List<ReferenceTableInfo> SetReferenceTables(List<IXLCell> tablesCells)
            {
                List<ReferenceTableInfo> referenceTables = new List<ReferenceTableInfo>();
                foreach (var tableCell in tablesCells)
                {
                    ReferenceTableInfo tableInfo = new ReferenceTableInfo();
                    tableInfo.TableCell = tableCell;
                    tableInfo.Queue = (int)tableCell.CellRight().Value;
                    tableInfo.Type = TableInfoBase.TableType.Reference;
                    tableInfo.TableName = tableCell.Value.ToString();
                    tableInfo.DoublesColumn = tableCell.CellRight(2).Value.ToString();
                    tableInfo.IdUpdateColumn = tableCell.CellRight(3).Value.ToString();
                    tableInfo.Worksheet = Workbook.Worksheets.Single(x => x.Name == tableCell.Value.ToString());
                    tableInfo.RowsUsedCount = tableInfo.Worksheet.RowsUsed().Count();

                    referenceTables.Add(tableInfo);
                }
                return referenceTables;
            }
        }
        public static void Dispose()
        {
            Workbook.Dispose();
            CommonTablesCells.Clear();
            MainTablesCells.Clear();
            CommonTablesInfos.Clear();
            MainTablesInfos.Clear();
        }
    }
}
