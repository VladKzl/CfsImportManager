using CfsImportManager;
using CfsImportManager.TablesInfo;

string excelPath = "D:\\ascon_obmen\\kozlov_vi\\Учет орг.техники\\CfsImportManager\\Склад 18 КВИ 090323 v22.03.2023.xlsx";
//DATABASE_URL=postgres://{user}:{password}@{hostname}:{port}/{database-name}
//Server=host;Port=5432;User Id=username;Password=secret;Database=databasename;
//Server=127.0.0.1;Port=82;Database=severcard;User Id=scuser;Password=123456;
string cfsConnectionString = "Server=192.168.0.112;Port=82;Database=severcart;User Id=scuser;Password=new_password2";

Console.WriteLine("Первичная компиляция проекта. Подождите..");
ExcelBase excelBse = new ExcelBase(excelPath);
Console.WriteLine("Первичная компиляция завершена.\n");

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
    bool continuation = false;
    ExcelBase.CommonTablesCells.ForEach(x => Console.WriteLine(x.Value));
    do
    {
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string manualTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.CommonTablesInfos.Single(x => x.TableName == manualTableName);
        Cfs cfs = new Cfs();
        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую","Завершить");
    }
    while (continuation);
}
void FillMainTables()
{
    bool continuation = false;
    ExcelBase.MainTablesCells.ForEach(x => Console.WriteLine(x.Value));
    do
    {
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string manualTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.MainTablesInfos.Single(x => x.TableName == manualTableName);
        Cfs cfs = new Cfs();
        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую", "Завершить");
    }
    while (continuation);
}
void FillReferenceTables()
{
    Console.WriteLine("");
    bool continuation = false;
    ExcelBase.MainTablesCells.ForEach(x => Console.WriteLine(x.Value));
    Console.WriteLine("Введите имя основной таблицы, справочники которой будем загружать\n");
    string mainTableName = Console.ReadLine();
    ExcelBase.MainTablesInfos.Single(x => x.TableName == mainTableName).ReferenceTables.ForEach(x => Console.WriteLine(x.TableName));
    do
    {
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string referencelTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.MainTablesInfos.Single(x => x.TableName == mainTableName).ReferenceTables.Single(x => x.TableName == referencelTableName);
        Cfs cfs = new Cfs();
        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую", "Завершить");
    }
    while (continuation);
}



