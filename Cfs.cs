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
                    try
                    {
                        DataAdapter.Update(DataTable);
                    }
                    catch (DBConcurrencyException)
                    {
                        DataAdapter.Fill(DataTable);
                        DataAdapter.Update(DataTable);
                    }
                    catch(Exception)
                    {
                        DataAdapter.Fill(DataTable);
                    }
/*                    DataAdapter.Update(DataTable);
                    DataAdapter.Fill(DataTable);*/
                    DataTable.AcceptChanges();
                }
                else
                {
                    AddNewRow();
/*                    DataAdapter.Update(DataTable);
                    DataTable.AcceptChanges();*/
                }
                CommonCode.GetPercent(i, tableInfo.RowsUsedCount);

                void FindeDoubles()
                {
                    string searchedValue = tableInfo.GetRowsCellFromTrimmed(i, tableInfo.DoublesColumn).CachedValue.ToString();

                    List<DataRow> existingRows;
                    IsColumnValueExists(tableInfo.DoublesColumn, searchedValue, out existingRows);

                    if (tableInfo.TableName == "ce_computer")
                        For_ce_computer();
                    else
                        OtherTables();

                    void For_ce_computer()
                    {
                        if (existingRows.Count == 0)
                        {
                            AddNewRow();
                            DoublesCounter.Add((searchedValue, 0));
                        }
                        if (existingRows.Count == 1)
                        {
                            if (!DoublesCounter.Any(x => x.doubledValue == searchedValue))
                            {
                                FillRow(i, existingRows[0]);
                                DoublesCounter.Add((searchedValue, 0));
                            }
                        }
                        if (existingRows.Count > 1)
                        {
                            if (!DoublesCounter.Any(x => x.doubledValue == searchedValue))
                            {
                                List<(DataRow row, long ticst)> rowAndTics = new List<(DataRow row, long ticsdat)>();
                                
                                for (int i_ = 0; i_ < existingRows.Count; i_++)
                                {
                                    rowAndTics.Add((existingRows[i_], existingRows[i_]["date_added"] is DBNull ? 0 : existingRows[i_].Field<DateTime>("date_added").Ticks));
                                }
                                long maxTics = rowAndTics.Select(x => x.ticst).Max(); // Самую позднюю

                                DataRow lastAddedRow = rowAndTics.Where(x => x.ticst == maxTics).Select(x => x.row).ToList().Last(); //Последнюю из самых поздних

                                existingRows.Remove(lastAddedRow);
                                existingRows.ForEach(x => x.SetField<int>("status", 1));

                                //Перезаписываем строку
                                FillRow(i, lastAddedRow);
                                DoublesCounter.Add((searchedValue, 0));
                            }
                        }
                    }
                    void OtherTables()
                    {
                        if (existingRows.Count == 0)
                        {
                            AddNewRow();
                        }
                        if (existingRows.Count == 1)
                        {
                            if (DoublesCounter.Any(x => x.doubledValue == searchedValue))
                            {
                                AddNewRow();
                            }
                            else
                            {
                                FillRow(i, existingRows[0]);
                                DoublesCounter.Add((searchedValue, 0));
                            }
                        }
                        if (existingRows.Count > 1)
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
                                /*tableInfo.DoublesColumn = string.Empty;*/
                            }
                        }
                    }
                }
                void AddNewRow()
                {
                    var newRow = DataTable.NewRow();
                    FillRow(i, newRow);
                    DataTable.Rows.Add(newRow);
                }
            }
            if (tableInfo.TableName == "ce_computer")
                ChangeStatusNonSortedToApplyed();
            DataAdapter.Update(DataTable);
            DataTable.AcceptChanges();

            void ChangeStatusNonSortedToApplyed()
            {
                DataTable.AsEnumerable().Where(x => x.Field<int>("status") == 7).ToList().ForEach(x => x.SetField<int>("status", 2));
            }
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
            Console.WriteLine("Начали синхронизацию экселя с db");
            Rebuild(tableInfo.TableName, cfsConnectionString);
            DoublesCounter.Clear();
            tableInfo.DefaultWorksheet = ExcelBase.WorkbookDefault.Worksheets.Single(x => x.Name == tableInfo.TableName);

            for (int i = 1; i < tableInfo.RowsUsedCount; i++)
            {
                string searchedValue = tableInfo.GetRowsCellFromDefault(i, tableInfo.IdUpdateColumn).CachedValue.ToString();

                List<DataRow> existingRows;
                if (!IsColumnValueExists(tableInfo.IdUpdateColumn, searchedValue, out existingRows))
                {
                    Console.WriteLine($"При обновлении excel не нашли совпадения по столбцу {tableInfo.IdUpdateColumn}. Такого не должно быть");
                    break;
                }
                if (tableInfo.TableName == "ce_computer")
                    Update_ce_computer();
                else
                    UpdateOtherTables();

                CommonCode.GetPercent(i, tableInfo.RowsUsedCount);

                void Update_ce_computer()
                {
                    DataRow oldesRow = existingRows.FirstOrDefault(x => x.Field<int>("status") == 2);
                    if (DoublesCounter.Any(x => x.doubledValue == searchedValue))
                        return;
                    if (oldesRow == null)
                        return;

                    int dbRowId = oldesRow.Field<int>("id");
                    tableInfo.GetRowsCellFromDefault(i, "id").Value = dbRowId;

                    DoublesCounter.Add((searchedValue, 0));
                }
                void UpdateOtherTables()
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
                    catch { }
                }
            }
            ExcelBase.WorkbookDefault.Save();
            Dispose();
        }
    }
}
