using CfsImportManager.TablesInfo;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CfsImportManager
{
    public class Cfs : CfsBase
    {
        public List<(string doubledValue, int nextValue)> DoublesCounter { get; set; } = new List<(string doubledValue, int nextValue)>();
        public void FillDbTable(TableInfoBase tableInfo, string cfsConnectionString)
        {
            Build(tableInfo.TableName, cfsConnectionString);

            for (int i = 1; i < tableInfo.RowsUsedCount; i++)
            {
                if(tableInfo.DoublesColumn != string.Empty)
                {
                    FindeDoubles();
                }
                else
                {
                    AddNewRow();
                }
                DataAdapter.Update(DataTable);
                DataTable.AcceptChanges();

                void FindeDoubles()
                {
                    string searchedValue = tableInfo.GetRowsCellFromTrimmed(i, tableInfo.DoublesColumn).CachedValue.ToString();

                    List<DataRow> existingRows;
                    if (IsColumnValueExists(tableInfo.DoublesColumn, searchedValue, out existingRows))
                    {
                        if (!DoublesCounter.Any(x => x.doubledValue == searchedValue))
                            DoublesCounter.Add((searchedValue, 0));
                        int nextRow = DoublesCounter.Single(x => x.doubledValue == searchedValue).nextValue;
                        try
                        {
                            FillRow(i, existingRows[nextRow]);
                            DoublesCounter.Remove((searchedValue, nextRow));
                            DoublesCounter.Add((searchedValue, nextRow + 1));
                        }
                        catch
                        {
                            AddNewRow();
                            tableInfo.DoublesColumn = string.Empty;
                        }
                    }
                    else
                    {
                        AddNewRow();
                    }
                }
                void AddNewRow()
                {
                    var newRow = DataTable.NewRow();
                    FillRow(i, newRow);
                    DataTable.Rows.Add(newRow);
                }
            }
/*            DataAdapter.Update(DataTable);
            DataTable.AcceptChanges();*/

            void FillRow(int i, DataRow row)
            {
                foreach (DataColumn column in DataTable.Columns)
                {
                    string columnName = column.ColumnName;
                    if (!tableInfo.IsColumnExistsFromTrimmed(columnName))
                    {
                        /*Console.WriteLine($"Колонка {columnName} таблицы {tableInfo.TableName} не найдена в той же таблице в ДБ");*/
                        continue;
                    }
                    var cellValue = tableInfo.GetRowsCellFromTrimmed(i, columnName).CachedValue;
                    if (cellValue.IsText)
                    {
                        row.SetField(columnName, cellValue.ToString());
                    }
                    if (cellValue.IsNumber)
                    {
                        row.SetField(columnName, (int)cellValue);
                    }
                    if (cellValue.IsDateTime)
                    {
                        row.SetField(columnName, (DateTime)cellValue);
                    }
                    if (cellValue.IsBlank)
                    {
                        row[columnName] = DBNull.Value;
                    }
                }
            }
        }
        public void UpdateExcelTable(TableInfoBase tableInfo, string cfsConnectionString, string excelPath)
        {
            Rebuild(tableInfo.TableName, cfsConnectionString);
            DoublesCounter.Clear();
            tableInfo.DefaultWorksheet = ExcelBase.WorkbookDefault.Worksheets.Single(x => x.Name == tableInfo.TableName);

            for (int i = 1; i < tableInfo.RowsUsedCount; i++)
            {
                string searchedValue = tableInfo.GetRowsCellFromDefault(i, tableInfo.IdUpdateColumn).CachedValue.ToString();

                List<DataRow> existingRows;
                if (IsColumnValueExists(tableInfo.IdUpdateColumn, searchedValue, out existingRows))
                {
                    if (!DoublesCounter.Any(x => x.doubledValue == searchedValue))
                        DoublesCounter.Add((searchedValue, 0));
                    int nextRow = DoublesCounter.Single(x => x.doubledValue == searchedValue).nextValue;
                    try
                    {
                        int dbRowId = existingRows[nextRow].Field<int>("id");
                        tableInfo.GetRowsCellFromDefault(i, "id").Value = dbRowId;

                        DoublesCounter.Remove((searchedValue, nextRow));
                        DoublesCounter.Add((searchedValue, nextRow + 1));
                    }
                    catch {}
                }
                else
                {
                    Console.WriteLine($"При обновлении excel не нашли совпадения по столбцу {tableInfo.IdUpdateColumn}. Такого не должно быть");
                }
            }
            ExcelBase.WorkbookDefault.Save();
            Dispose();
        }
    }
}
