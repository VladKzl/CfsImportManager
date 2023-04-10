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
        public enum TableType
        {
            Main,
            Common,
            Reference
        }
        public IXLCell TableCell { get; set; }
        public bool DownloadedColorStatus
        {
            get
            {
                if(TableCell.Style.Fill.BackgroundColor != XLColor.AppleGreen)
                    return false;
                return true;
            }
            set
            {
                if(value == true)
                    TableCell.Style.Fill.BackgroundColor = XLColor.AppleGreen;
                if (value == false)
                    TableCell.Style.Fill.BackgroundColor = XLColor.NoColor;
            }
        }
        public int Queue { get; set; }
        public TableType Type { get; set; }
        public string DoublesColumn { get; set; }
        public string IdUpdateColumn { get; set; }
        public string TableName { get; set; }
        public int RowsUsedCount { get; set; }
        public IXLWorksheet Worksheet { get; set; }
        public IXLCell GetXLCellFromRow(int rowIndex, string columnName)
        {
            int lastCell = Worksheet.ColumnsUsed().Count();
            return Worksheet.RowsUsed().ElementAt(rowIndex).Cells(1, lastCell).Single(x => (string)x.WorksheetColumn().Cell(1).Value == columnName);
        }
        public IXLRow GetXLRow(string columnName, string uiniqueCellValue)
        {
            var column = Worksheet.Row(1).Cells().Single(x => (string)x.Value == columnName).WorksheetColumn();
            return column.CellsUsed().Single(x => (string)x.Value == uiniqueCellValue).WorksheetRow();
        }
        public bool IsXLColumnExists(string columnName)
        {
            int lastCell = Worksheet.ColumnsUsed().Count();
            return Worksheet.Row(1).Cells(1, lastCell).Any(x => (string)x.Value == columnName);
        }
        public bool IsSealed(IXLCell cell)
        {
            //fontName
            //ff92d050 - зеленый
            //ffb2b2b2 - серый
            //Логика усложнена так как цвета глючат.
            try
            {
                if (cell.WorksheetColumn().Cell(1).Style.Fill.BackgroundColor.Color.Name != "ff92d050" 
                                           && cell.Style.Fill.BackgroundColor.Color.Name != "ffb2b2b2")
                    return true;
            }
            catch
            {
                return true;
            }
            return false;
        }
    }
}
