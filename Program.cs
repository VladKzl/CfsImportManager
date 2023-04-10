using CfsImportManager;
using CfsImportManager.TablesInfo;

string excelPath = "B:\\ДИИТ\\Severcart\\Склад 18 КВИ актуальный.xlsx";
/*string excelPath = "D:\\ascon_obmen\\kozlov_vi\\Учет орг.техники\\CfsImportManager\\Склад 18 КВИ актуальный.xlsx";*/
string cfsConnectionString = "Server=192.168.0.112;Port=82;Database=severcart;User Id=scuser;Password=new_password2";

InitializeExcel();
Console.WriteLine
(
    "Выберите, какие тублицы грузим:\n" +
    "1 - Общие таблицы\n" +
    "2 - Основные таблицы\n" +
    "3 - Справочные таблицы\n" +
    "4 - Частично автоматически\n"
);

int workModeNum;
int.TryParse(Console.ReadLine(), out workModeNum);

switch(workModeNum)
{
    case 1:FillCommonTables(); break;
    case 2:FillMainTables(); break;
    case 3:FillReferenceTables(); break;
    case 4:FillManualAuto(); break;
}
void FillManualAuto()
{
    Cfs cfs = new Cfs();
    if (!FillCommonTables())
        return;

    Console.WriteLine($"Переходим к загрузке главных таблиц и их справочников");
    var downloaded = ExcelBase.MainTablesInfos.Where(x => x.DownloadedColorStatus == true).ToList();
    downloaded.ForEach(x => Console.WriteLine($"{x.TableName} - загружена"));
    var notDownloaded = ExcelBase.MainTablesInfos.Where(x => x.DownloadedColorStatus == true).ToList();
    notDownloaded.ForEach(x => Console.WriteLine($"{x.TableName}"));
    while (true)
    {
        if (!ExcelBase.MainTablesInfos.Any(x => x.DownloadedColorStatus == false))
            break;
        Console.WriteLine("Введите имя таблицы:\n");
        string mainTableName = Console.ReadLine();
        MainTableInfo tableInfo = ExcelBase.MainTablesInfos.Single(x => x.TableName == mainTableName);
        if (FillMainTable(tableInfo))
            break;
    }

    bool FillCommonTables()
    {
        for (int i = 0; i < ExcelBase.CommonTablesInfos.Count(); i++)
        {
            var tableInfo = ExcelBase.CommonTablesInfos.Single(x => x.Queue == i + 1);
            if (tableInfo.DownloadedColorStatus)
            {
                Console.WriteLine($"{i + 1}|{tableInfo.TableName} - загружена");
                continue;
            }
            Console.WriteLine($"{i + 1}|{tableInfo.TableName} - следующая к загрузке. Загружаем?");
            if(!CommonCode.UserValidationPlusOrMinus("+", "-"))
                return false;

            cfs.FillDbTable(tableInfo, cfsConnectionString);
            cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);
            tableInfo.DownloadedColorStatus = true;
        }
        return true;
    }
    bool FillMainTable(MainTableInfo tableInfo)
    {
        if (!DownloadReferenceTables())
            return false;

        Console.WriteLine($"{tableInfo.Queue}|{tableInfo.TableName} - следующая к загрузке. Загружаем?");
        if (!CommonCode.UserValidationPlusOrMinus("+", "-"))
            return false;

        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);
        tableInfo.DownloadedColorStatus = true;

        return true;

        
        bool DownloadReferenceTables()
        {
            for (int i = 0; i < tableInfo.ReferenceTables.Count(); i++)
            {
                var refTable = tableInfo.ReferenceTables.Single(x => x.Queue == i + 1);
                if (refTable.DownloadedColorStatus)
                {
                    Console.WriteLine($"{i + 1}|{refTable.TableName} - загружена");
                    continue;
                }
                Console.WriteLine($"{i + 1}|{refTable.TableName} - следующая к загрузке. Загружаем?");
                if (!CommonCode.UserValidationPlusOrMinus("+", "-"))
                    return false;

                cfs.FillDbTable(refTable, cfsConnectionString);
                cfs.UpdateExcelTable(refTable, cfsConnectionString, excelPath);
                refTable.DownloadedColorStatus = true;
            }
            return true;
        }
    }
}
void FillCommonTables()
{
    Cfs cfs = new Cfs();
    bool continuation = true;
    bool persentation = false;
    do
    {
        /*InitializeExcel();*/
        ExcelBase.Workbook.RecalculateAllFormulas();
        if (!persentation)
        {
            ExcelBase.CommonTablesCells.ForEach(x => Console.WriteLine(x.Value));
            persentation = true;
        }
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string manualTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.CommonTablesInfos.Single(x => x.TableName == manualTableName);
        
        var test = tableInfo.Worksheet.Row(5).Cells().ToList().Select(x => x.CachedValue).ToList();//

        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        /*ExcelBase.Dispose();*/
        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую","Завершить");
    }
    while (continuation);
}
void FillMainTables()
{
    Cfs cfs = new Cfs();
    bool continuation = true;
    bool persentation = false;
    do
    {
        /*InitializeExcel();*/
        ExcelBase.Workbook.RecalculateAllFormulas();
        if (!persentation)
        {
            ExcelBase.MainTablesCells.ForEach(x => Console.WriteLine(x.Value));
            persentation = true;
        }
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string manualTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.MainTablesInfos.Single(x => x.TableName == manualTableName);

        var test = tableInfo.Worksheet.Row(5).Cells().ToList().Select(x => x.CachedValue).ToList();

        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        /*ExcelBase.Dispose();*/
        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую", "Завершить");
    }
    while (continuation);
}
void FillReferenceTables()
{
    bool continuation = true;
    bool persentation = false;

    ExcelBase.MainTablesCells.ForEach(x => Console.WriteLine(x.Value));
    Console.WriteLine("Введите имя основной таблицы, справочники которой будем загружать\n");
    string mainTableName = Console.ReadLine();
    do
    {
        ExcelBase.Workbook.RecalculateAllFormulas();

        if (!persentation)
        {
            ExcelBase.MainTablesInfos.Single(x => x.TableName == mainTableName).ReferenceTables.ForEach(x => Console.WriteLine(x.TableName));
            persentation = true;
        }
        Console.WriteLine("Введите имя загружаемой таблицы\n");
        string referencelTableName = Console.ReadLine();
        TableInfoBase tableInfo = ExcelBase.MainTablesInfos.Single(x => x.TableName == mainTableName).ReferenceTables.Single(x => x.TableName == referencelTableName);
        Cfs cfs = new Cfs();

        var test2 = tableInfo.Worksheet.Row(5).Cells().ToList().Select(x => x.CachedValue).ToList();

        cfs.FillDbTable(tableInfo, cfsConnectionString);
        cfs.UpdateExcelTable(tableInfo, cfsConnectionString, excelPath);

        Console.WriteLine("Загрузили.\n");
        continuation = CommonCode.UserValidationPlusOrMinus("Взять следующую", "Завершить");
/*        if(continuation)
            InitializeExcel();*/
    }
    while (continuation);
}
void InitializeExcel()
{
    Console.WriteLine("Получаем изменения из экселя..\n");
    new ExcelBase(excelPath);
}



