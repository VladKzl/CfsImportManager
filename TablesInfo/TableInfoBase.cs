using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CfsImportManager.TablesInfo
{
    public class TableInfoBase
    {
        public int Queue { get; set; }
        public string DoublesColumn { get; set; }
        public string IdUpdateColumn { get; set; }
        public string TableName { get; set; }
        public int RowsUsedCount { get; set; }
        public IXLWorksheet DefaultWorksheet { get; set; }
        public IXLWorksheet TrimmedWorksheet { get; set; }
        public IXLCell GetRowsCellFromTrimmed(int rowIndex, string columnName)
        {
            int lastCell = TrimmedWorksheet.ColumnsUsed().Count();
            return TrimmedWorksheet.RowsUsed().ElementAt(rowIndex).Cells(1, lastCell).Single(x => (string)x.WorksheetColumn().Cell(1).Value == columnName);
        }
        public IXLRow GetRowFromTrimmed(string columnName, string uiniqueCellValue)
        {
            var column = TrimmedWorksheet.Row(1).Cells().Single(x => (string)x.Value == columnName).WorksheetColumn();
            return column.CellsUsed().Single(x => (string)x.Value == uiniqueCellValue).WorksheetRow();
        }
        public bool IsColumnExistsFromTrimmed(string columnName)
        {
            int lastCell = TrimmedWorksheet.ColumnsUsed().Count();
            return TrimmedWorksheet.Row(1).Cells(1, lastCell).Any(x => (string)x.Value == columnName);
        }
        public IXLCell GetRowsCellFromDefault(int rowIndex, string columnName)
        {
            int lastCell = DefaultWorksheet.ColumnsUsed().Count();
            return DefaultWorksheet.RowsUsed().ElementAt(rowIndex).Cells(1, lastCell).Single(x => (string)x.WorksheetColumn().Cell(1).Value == columnName);
        }
        public IXLRow GetRowFromDefault(string columnName, string uiniqueCellValue)
        {
            var column = DefaultWorksheet.Row(1).Cells().Single(x => (string)x.Value == columnName).WorksheetColumn();
            return column.CellsUsed().Single(x => (string)x.Value == uiniqueCellValue).WorksheetRow();
        }
        public bool IsColumnExistsFromDefault(string columnName)
        {
            int lastCell = DefaultWorksheet.ColumnsUsed().Count();
            return DefaultWorksheet.Row(1).Cells(1, lastCell).Any(x => (string)x.Value == columnName);
        }
    }
}
