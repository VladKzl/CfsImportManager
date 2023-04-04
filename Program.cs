using CfsImportManager;
using CfsImportManager.TablesInfo;

string excelPath = "B:\\ДИИТ\\Severcart\\Склад 18 КВИ v030423.xlsx";
string cfsConnectionString = "Server=192.168.0.112;Port=82;Database=severcart;User Id=scuser;Password=new_password2";
//DATABASE_URL=postgres://{user}:{password}@{hostname}:{port}/{database-name}
//Server=host;Port=5432;User Id=username;Password=secret;Database=databasename;
//Server=127.0.0.1;Port=82;Database=severcard;User Id=scuser;Password=123456;


Console.WriteLine
(
    "Выберите, какие тублицы грузим:\n" +
    "1 - Общие таблицы\n" +
    "2 - Основные таблицы\n" +
    "3 - Справочные таблицы"
);

int workModeNum;
int.TryParse(Console.ReadLine(), out workModeNum);

switch(workModeNum)
{
    case 1:FillCommonTables(); break;
    case 2:FillMainTables(); break;
    case 3:FillReferenceTables(); break;
}
void FillCommonTables()
{
    bool continuation = true;
    bool persentation = false;
    do
    {
        InitializeExcel();
        if (!persentation)
        {
            ExcelBase.CommonTablesCells.ForEach(x => Console.WriteLine(x.Value));
            persentation = true;
        }
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string manualTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.CommonTablesInfos.Single(x => x.TableName == manualTableName);
        Cfs cfs = new Cfs();
        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        ExcelBase.Dispose();
        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую","Завершить");
    }
    while (continuation);
}
void FillMainTables()
{
    bool continuation = true;
    bool persentation = false;
    do
    {
        InitializeExcel();
        if (!persentation)
        {
            ExcelBase.MainTablesCells.ForEach(x => Console.WriteLine(x.Value));
            persentation = true;
        }
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string manualTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.MainTablesInfos.Single(x => x.TableName == manualTableName);
        Cfs cfs = new Cfs();
        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        ExcelBase.Dispose();
        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую", "Завершить");
    }
    while (continuation);
}
void FillReferenceTables()
{
    bool continuation = true;
    bool persentation = false;
    InitializeExcel();
    ExcelBase.MainTablesCells.ForEach(x => Console.WriteLine(x.Value));
    Console.WriteLine("Введите имя основной таблицы, справочники которой будем загружать\n");
    string mainTableName = Console.ReadLine();
    do
    {
        if (!persentation)
        {
            ExcelBase.MainTablesInfos.Single(x => x.TableName == mainTableName).ReferenceTables.ForEach(x => Console.WriteLine(x.TableName));
            persentation = true;
        }
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string referencelTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.MainTablesInfos.Single(x => x.TableName == mainTableName).ReferenceTables.Single(x => x.TableName == referencelTableName);
        Cfs cfs = new Cfs();
        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        ExcelBase.Dispose();
        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую", "Завершить");
        if(continuation)
            InitializeExcel();
    }
    while (continuation);
}
void InitializeExcel()
{
    Console.WriteLine("Получаем изменения из экселя..\n");
    new ExcelBase(excelPath);
}



