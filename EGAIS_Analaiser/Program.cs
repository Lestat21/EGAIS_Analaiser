using EGAIS_Analaiser.Data;
using EGAIS_Analaiser.ParserXLSX;
using EGAIS_Analaiser.View;
using Newtonsoft.Json;

Console.WriteLine("======= Приступим =======");
Console.WriteLine($"{DateTime.Now} - Очищаем таблицы от стрых данных");

using (UserContext dbContext = new()) // чистим все таблицы
{
    dbContext.Remains1Cs.RemoveRange(dbContext.Remains1Cs);
    dbContext.Remains.RemoveRange(dbContext.Remains);
    dbContext.Sellings.RemoveRange(dbContext.Sellings);
    dbContext.FullShort1Cs.RemoveRange(dbContext.FullShort1Cs);
    dbContext.Zagotovkas.RemoveRange(dbContext.Zagotovkas);
    dbContext.Sklads.RemoveRange(dbContext.Sklads);
    dbContext.TDLes.RemoveRange(dbContext.TDLes);
    dbContext.FullShort1Cs.RemoveRange(dbContext.FullShort1Cs);
    dbContext.SaveChanges();
}

try
{
    Console.WriteLine($"{DateTime.Now} - Загружаем пути к файлам данных");
    var configFile = "config.json";

    if (!File.Exists(configFile))
    {
        throw new FileNotFoundException("Файл конфигурации не найден", configFile);
    }

    var json = File.ReadAllText(configFile);

    Config config;
    try
    {
        config = JsonConvert.DeserializeObject<Config>(json);
    }
    catch (JsonException ex)
    {
        throw new InvalidDataException("Файл конфигурации некорректен", ex);
    }

    var directoryPath = config.DirectoryPath;
    var filePath = config.FilePath;
    var filePathSell = config.FilePathSell;
    var filePathZag = config.FilePathZag;
    var filePathSklad = config.FilePathSklad;
    var filePathTDLes = config.FilePathTDLes;

    var dir = "d:\\EGAIS\\";
    var outFile = "d:\\EGAIS\\Отчет ЕГАИС на " + DateTime.Now.ToString("yyyy.MM.dd HH.mm.ss") + ".xlsx";
    if (!Directory.Exists(dir))
    {
        throw new DirectoryNotFoundException("Директория выгрузки файла не существует");
    }
       
    Console.WriteLine($"{DateTime.Now} - Собираем остатки по складам из ЕГАИС");
    ParserEGAIS.ParserEgais(filePath);              // в параметры ссылка на файл собираем остатки ЕГАИС
    Console.WriteLine($"{DateTime.Now} - Собираем данные о расходе из ЕГАИС");
    ParserEGAIS.ParserSellingEGAIS(filePathSell);   // собираем реализацию ЕГАИС
    Console.WriteLine($"{DateTime.Now} - Собираем данные о заготовке из ЕГАИС");
    ParserEGAIS.ParserZagotovkaEGAIS(filePathZag);  // собираем заготовку (оперативный учет)
    Console.WriteLine($"{DateTime.Now} - Создаем список складов ЕГАИС");
    ParserEGAIS.ParserSkladEGAIS(filePathSklad);    // собираем список складов с ЕГАИС
    Console.WriteLine($"{DateTime.Now} - Собираем данные о ТД-Лес из ЕГАИС");
    ParserEGAIS.ParserTDLes(filePathTDLes);         // собираем данные по ТД Лес

    Console.WriteLine($"{DateTime.Now} - Собираем данные обобщенные из 1С");
    Parser1C.FullShort1C(directoryPath);

    //RemainsToConsole.RemainsToCW();               // выводим в консоль
    //SellingsToConsole.SellingToCW();              // реализация на консоль
    //ZogotovkaToConsole.ZagotovkaToCW();           // оперативный на консоль

    SaveToFile.ToFile(outFile);
    Console.WriteLine($"{DateTime.Now} - Все готово. Выходной файл лежит в {outFile}\nМожете закрыть программу нажав любую клавишу.");
}
catch (Exception ex)
{
    Console.WriteLine($"Произошла ошибка: Файл поврежден или не найден. {ex.Message}");
}
Console.ReadKey();