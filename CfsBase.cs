using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CfsImportManager
{
    public class CfsBase
    {
        public NpgsqlConnection Connection { get; set; } = new NpgsqlConnection();
        public NpgsqlDataAdapter DataAdapter { get; set; } = new NpgsqlDataAdapter();
        public NpgsqlCommandBuilder SqlCommandBuilder { get; set; } = new NpgsqlCommandBuilder();
        public DataTable DataTable { get; set; } = new DataTable();
        public void Build(string tableName, string cfsConnectionString)
        {
            DataTable.TableName = tableName;
            Connection = new NpgsqlConnection(cfsConnectionString);
            DataAdapter = new NpgsqlDataAdapter($"select * from {tableName}", Connection);
            /*DataAdapter.MissingMappingAction = MissingMappingAction.Error;
            DataAdapter.MissingSchemaAction = MissingSchemaAction.Error;*/

            SqlCommandBuilder = new NpgsqlCommandBuilder(DataAdapter);
            DataAdapter.DeleteCommand = SqlCommandBuilder.GetDeleteCommand(true);
            DataAdapter.UpdateCommand = SqlCommandBuilder.GetUpdateCommand(true);
            DataAdapter.InsertCommand = SqlCommandBuilder.GetInsertCommand(true);

            DataAdapter.RowUpdating += new NpgsqlRowUpdatingEventHandler(OnRowUpdating);
            DataAdapter.RowUpdated += new NpgsqlRowUpdatedEventHandler(OnRowUpdated);

            DataAdapter.FillSchema(DataTable, SchemaType.Source);
            DataAdapter.Fill(DataTable);
        }
        public void Rebuild(string tableName, string cfsConnectionString)
        {
            Dispose();
            Build(tableName, cfsConnectionString);
        }
        public bool IsColumnValueExists(string mainColumn, string searchedValue, out List<DataRow> existingRow)
        {
            existingRow = DataTable.AsEnumerable().Where(x => x.Field<string>(mainColumn) == searchedValue).ToList();
            if (existingRow != null & existingRow.Count() != 0)
                return true;
            return false;
        }
        public void Dispose()
        {
            if (Connection.State != ConnectionState.Closed)
                Connection.Close();
            Connection.Dispose();
            DataTable.Clear();
            DataTable.Dispose();
            DataAdapter.Dispose();
            DataAdapter.Dispose();
            SqlCommandBuilder.Dispose();
        }
        public void OnRowUpdating(object sender, NpgsqlRowUpdatingEventArgs args)
        {
            if (args.Status == UpdateStatus.ErrorsOccurred)
            {
                args.Row.RowError = args.Errors.Message;
                args.Status = UpdateStatus.SkipCurrentRow;
            }
        }
        public void OnRowUpdated(object sender, NpgsqlRowUpdatedEventArgs args)
        {
            if (args.Status == UpdateStatus.ErrorsOccurred)
            {
                args.Row.RowError = args.Errors.Message;
                args.Status = UpdateStatus.SkipCurrentRow;
            }
        }
    }
}
