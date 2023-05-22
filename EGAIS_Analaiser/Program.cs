using EGAIS_Analaiser.Data;
using EGAIS_Analaiser.ParserXLSX;
using EGAIS_Analaiser.View;

using (var dbContext = new UserContext()) //чистим все таблицы
{
    // Удалить все существующие записи
    dbContext.Remains1Cs.RemoveRange(dbContext.Remains1Cs);
    dbContext.Remains.RemoveRange(dbContext.Remains);
    dbContext.Sellings.RemoveRange(dbContext.Sellings);
    dbContext.FullShort1Cs.RemoveRange(dbContext.FullShort1Cs);
    dbContext.Zagotovkas.RemoveRange(dbContext.Zagotovkas);
    dbContext.Sklads.RemoveRange(dbContext.Sklads);
    dbContext.TDLes.RemoveRange(dbContext.TDLes);
    dbContext.SaveChanges();
}

// задаем пути для файлов
var directoryPath = "d:\\EGAIS\\1C\\1c.xlsx";


var filePath = "d:\\EGAIS\\EGAIS\\reamains.xlsx";
var filePathSell = "d:\\EGAIS\\EGAIS\\SellingEGAIS.xlsx";
var filePathZag = "d:\\EGAIS\\EGAIS\\Zagotovka.xlsx";
var filePathSklad = "d:\\EGAIS\\EGAIS\\SkladySpisok.xlsx";
var filePathTDLes = "d:\\EGAIS\\EGAIS\\TDLes.xlsx";

var outFile = "d:\\EGAIS\\Отчет ЕГАИС на " + DateTime.Now.ToString("yyyy.MM.dd HH.mm.ss") + ".xlsx";


ParserEGAIS.ParserEgais(filePath); // в параметры ссылка на файл собираем остатки ЕГАИС
ParserEGAIS.ParserSellingEGAIS(filePathSell); //собираем реализацию ЕГАИС
ParserEGAIS.ParserZagotovkaEGAIS(filePathZag); //собираем заготовку (оперативный учет)
ParserEGAIS.ParserSkladEGAIS(filePathSklad); // собираем список складов с ЕГАИС
ParserEGAIS.ParserTDLes(filePathTDLes); // собираем данные по ТД Лес

Parser1C.FullShort1C(directoryPath);

RemainsToConsole.RemainsToCW(); //выводим в консоль
SellingsToConsole.SellingToCW(); // реализация на консоль
ZogotovkaToConsole.ZagotovkaToCW(); // оперативный на консоль
SaveToFile.ToFile(outFile);


Console.ReadKey();

