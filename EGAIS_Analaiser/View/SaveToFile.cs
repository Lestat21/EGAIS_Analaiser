using EGAIS_Analaiser.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EGAIS_Analaiser.View
{
    public class SaveToFile
    {
        public static void ToFile(string newFilePath)
        {
            using (var dbContext = new UserContext())
            {
                // тут выбираем все необходимые данные из таблиц

                var resultsEGAIS = dbContext.Remains
                         .GroupBy(record => record.WarehouseOwner)
                         .Select(group => new
                         {
                             WarehouseOwner = group.Key,
                             TotalVolume = group.Sum(record => record.Volume)
                         })
                         .OrderBy(result => result.WarehouseOwner);

                var zagotovkaEGAIS = dbContext.Zagotovkas
                         .GroupBy(record => record.Forestry)
                         .Select(group => new
                         {
                             WarehouseOwner = group.Key,
                             TotalVolume = group.Sum(record => record.OperationalAccountingTotal)
                         })
                         .OrderBy(result => result.WarehouseOwner);

                var selling = dbContext.Sellings
                        .GroupBy(record => record.Division)
                        .Select(group => new
                        {
                            Division = group.Key,
                            TotalVolume = group.Sum(record => record.Volume)
                        })
                        .OrderBy(result => result.Division);

                var result1C = dbContext.FullShort1Cs
                          .Select(record => new
                          {
                              Subdivision = record.Subdivision,
                              Zagotovka = record.Procurement,
                              Balance = record.Balance,
                              Sale = record.Sale,
                              SelfConsumption = record.SelfConsumption,
                              Processing = record.Processing
                          }).OrderBy(result => result.Subdivision).ToList();

               


                // тут формируем книгу эксель.


                using (var package = new ExcelPackage(new FileInfo(newFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.Add("ОБЩАЯ АНАЛИТИКА");
                  
                    #region == Первый лист ОБЩАЯ АНАЛИТИКА

                    // Написать заголовки в шапку таблицы
                    worksheet.Cells[1, 2, 1, 5].Merge = true;
                    worksheet.Cells[1, 2].Value = "Остатки";
                    worksheet.Cells[1, 6, 1, 9].Merge = true;
                    worksheet.Cells[1, 6].Value = "Заготовка";
                    worksheet.Cells[1, 10, 1, 13].Merge = true;
                    worksheet.Cells[1, 10].Value = "Расход";

                    worksheet.Cells[1, 1, 2, 1].Merge = true;
                    worksheet.Cells[1, 1].Value = "Владелец склада";

                    worksheet.Cells[2, 2].Value = "ЕГАИС";
                    worksheet.Cells[2, 3].Value = "1C";
                    worksheet.Cells[2, 4].Value = "%";
                    worksheet.Cells[2, 5].Value = "+/- 1C";

                    worksheet.Cells[2, 6].Value = "ЕГАИС";
                    worksheet.Cells[2, 7].Value = "1С";
                    worksheet.Cells[2, 8].Value = "%";
                    worksheet.Cells[2, 9].Value = "+/- 1C";

                    worksheet.Cells[2, 10].Value = "ЕГАИС";
                    worksheet.Cells[2, 11].Value = "1С";
                    worksheet.Cells[2, 12].Value = "%";
                    worksheet.Cells[2, 13].Value = "+/- 1C";

                    worksheet.Cells[1, 14, 2, 14].Merge = true;
                    worksheet.Cells[3, 14, 14, 14].Merge = true;
                    worksheet.Cells[1, 14].Value = "ОБЩИЙ %";

                    int row = 3;

                    // Написать данные

                    foreach (var result in resultsEGAIS)
                    {
                        var r1c = result1C.FirstOrDefault(p => p.Subdivision == result.WarehouseOwner)?.Balance ?? 0;

                        worksheet.Cells[row, 1].Value = result.WarehouseOwner;
                        worksheet.Cells[row, 2].Value = result.TotalVolume;
                        worksheet.Cells[row, 3].Value = r1c;
                        worksheet.Cells[row, 4].Value = result.TotalVolume / r1c * 100;
                        worksheet.Cells[row, 5].Value = r1c - result.TotalVolume;
                        row++;
                    }

                    row = 3;
                    foreach (var result in zagotovkaEGAIS)
                    {
                        var r1c = result1C.FirstOrDefault(p => p.Subdivision == result.WarehouseOwner)?.Zagotovka ?? 0;

                        worksheet.Cells[row, 6].Value = result.TotalVolume;
                        worksheet.Cells[row, 7].Value = r1c;
                        worksheet.Cells[row, 8].Value = result.TotalVolume / r1c * 100;
                        worksheet.Cells[row, 9].Value = r1c - result.TotalVolume;
                        row++;
                    }

                    row = 3;
                    foreach (var result in selling)
                    {
                        var r1c = (result1C.FirstOrDefault(p => p.Subdivision == result.Division)?.Sale ?? 0) +
                            (result1C.FirstOrDefault(p => p.Subdivision == result.Division)?.SelfConsumption ?? 0) +
                            (result1C.FirstOrDefault(p => p.Subdivision == result.Division)?.Processing ?? 0);

                        worksheet.Cells[row, 10].Value = result.TotalVolume;
                        worksheet.Cells[row, 11].Value = r1c;
                        worksheet.Cells[row, 12].Value = result.TotalVolume / r1c * 100;
                        worksheet.Cells[row, 13].Value = r1c - result.TotalVolume;
                        row++;
                    }

                    // Написать итоги
                    worksheet.Cells[row, 1].Value = "";
                    worksheet.Cells[row, 2].Value = resultsEGAIS.Sum(r => r.TotalVolume);
                    worksheet.Cells[row, 3].Value = result1C.Sum(r => r.Balance);
                    worksheet.Cells[row, 4].Value = resultsEGAIS.Sum(r => r.TotalVolume) / result1C.Sum(r => r.Balance) * 100;
                    worksheet.Cells[row, 5].Value = result1C.Sum(r => r.Balance) - resultsEGAIS.Sum(r => r.TotalVolume);
                    worksheet.Cells[row, 6].Value = zagotovkaEGAIS.Sum(r => r.TotalVolume);
                    worksheet.Cells[row, 7].Value = result1C.Sum(r => r.Zagotovka);
                    worksheet.Cells[row, 8].Value = zagotovkaEGAIS.Sum(r => r.TotalVolume) / result1C.Sum(r => r.Zagotovka) * 100;
                    worksheet.Cells[row, 9].Value = result1C.Sum(r => r.Zagotovka) - zagotovkaEGAIS.Sum(r => r.TotalVolume);
                    worksheet.Cells[row, 10].Value = selling.Sum(r => r.TotalVolume);
                    worksheet.Cells[row, 11].Value = result1C.Sum(r => r.SelfConsumption) + result1C.Sum(r => r.Sale) + result1C.Sum(r => r.Processing);
                    worksheet.Cells[row, 12].Value = selling.Sum(r => r.TotalVolume) / (result1C.Sum(r => r.SelfConsumption) + result1C.Sum(r => r.Sale) + result1C.Sum(r => r.Processing)) * 100;
                    worksheet.Cells[row, 13].Value = result1C.Sum(r => r.SelfConsumption) + result1C.Sum(r => r.Sale) + result1C.Sum(r => r.Processing) - selling.Sum(r => r.TotalVolume);
                    worksheet.Cells[row, 14].Value = Math.Abs((resultsEGAIS.Sum(r => r.TotalVolume) / result1C.Sum(r => r.Balance) * 100) - 1) +
                        Math.Abs((zagotovkaEGAIS.Sum(r => r.TotalVolume) / (result1C.Sum(r => r.Zagotovka) * 100) - 1)) +
                        Math.Abs((selling.Sum(r => r.TotalVolume) / (result1C.Sum(r => r.SelfConsumption) + result1C.Sum(r => r.Sale) + result1C.Sum(r => r.Processing) * 100) - 1));
                    #endregion

                    #region  ==  форматирование первого листа
                    worksheet.Column(1).Width = 40;

                    for (int i = 2; i <= 14; i++)
                    {
                        worksheet.Column(i).Width = 10;
                    }

                    for (int column = 1; column <= 14; column++)
                    {
                        worksheet.Cells[1, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[1, column].Style.Font.Bold = true;
                        worksheet.Cells[2, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, column].Style.Font.Bold = true;
                        worksheet.Cells[15, column].Style.Font.Bold = true;
                    }

                    worksheet.Cells[3, 4, 15, 4].Style.Numberformat.Format = "0.00";
                    worksheet.Cells[3, 8, 15, 8].Style.Numberformat.Format = "0.00";
                    worksheet.Cells[3, 12, 15, 12].Style.Numberformat.Format = "0.00";
                    worksheet.Cells[15, 14].Style.Numberformat.Format = "0.00";

                    for (int i = 1; i <= 15; i++)
                    {
                        for (int column = 1; column <= 14; column++)
                        {
                            ExcelRange cell = worksheet.Cells[i, column];
                            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }
                    #endregion


                    var worksheet1 = package.Workbook.Worksheets.Add("ПУСТЫЕ СКЛАДЫ");

                    var skladEGAIS = dbContext.Sklads
                       .Select(record => new
                       {
                           Sklad = record.Name,
                           WarehouseOwner = record.WarehouseOwner,
                           CreateData = record.OpenDate
                       }).ToList();

                    var remainsBySlad = dbContext.Remains
                        .GroupBy(record => record.Warehouse)
                        .Select(group => new
                        {
                            Sklad = group.Key,
                            TotalVolume = group.Sum(record => record.Volume)
                        }).ToList();

                    var skladToRemove = skladEGAIS
                                        .Where(s => !remainsBySlad.Any(r => r.Sklad == s.Sklad))
                                        .OrderBy(s => s.WarehouseOwner)
                                        .ToList();

                    foreach (var sklad in skladToRemove)
                    {
                        skladEGAIS.Remove(sklad);
                    }

                    worksheet1.Column(1).Width = 30;
                    worksheet1.Column(2).Width = 70;
                    worksheet1.Column(3).Width = 15;
                    worksheet1.Cells[2, 1].Value = "Лесничество";
                    worksheet1.Cells[2, 2].Value = "Наименование склада";
                    worksheet1.Cells[2, 3].Value = "Дата открытия склада";

                    row = 3;
                    foreach (var sklad in skladToRemove)
                    {
                        worksheet1.Cells[row, 1].Value = sklad.WarehouseOwner;
                        worksheet1.Cells[row, 2].Value = sklad.Sklad;
                        worksheet1.Cells[row, 3].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet1.Cells[row, 3].Value = sklad.CreateData;
                        row++;
                    }

                    for (int i = 2; i <= row; i++)
                    {
                        for (int column = 1; column <= 3; column++)
                        {
                            ExcelRange cell = worksheet1.Cells[i, column];
                            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    var worksheet2 = package.Workbook.Worksheets.Add("Нарушения. Учет ТДЛЕС");

                    var resultTDLes = dbContext.TDLes
                        .Where(record => record.ServerProcessingDateTime > record.DocumentDate.AddDays(1) && record.DocumentType.Contains("Расход"))
                        .Select(record => new
                        {
                            WarehouseOwner = record.WarehouseOwner,
                            NumTDLes = record.DocumentNumber,
                            DocType = record.DocumentType,
                            DocumentDate = record.DocumentDate,
                            ServerProcessingDateTime = record.ServerProcessingDateTime,
                            Percon = record.CreatedByUser,
                            DaysDifference = (record.ServerProcessingDateTime - record.DocumentDate).TotalDays

                        })
                        .OrderBy(record => record.Percon).ToList();

                    worksheet2.Column(1).Width = 30;
                    worksheet2.Column(2).Width = 30;
                    worksheet2.Column(3).Width = 30;
                    worksheet2.Column(4).Width = 15;
                    worksheet2.Column(5).Width = 15;
                    worksheet2.Column(6).Width = 30;
                    worksheet2.Cells[2, 1].Value = "Лесничество";
                    worksheet2.Cells[2, 2].Value = "Номер документа";
                    worksheet2.Cells[2, 3].Value = "Тип документа";
                    worksheet2.Cells[2, 4].Value = "Дата документа";
                    worksheet2.Cells[2, 5].Value = "Время сервера";
                    worksheet2.Cells[2, 6].Value = "Пользователь";   
                    worksheet2.Cells[2, 7].Value = "Расхождение";

                    row = 3;
                    foreach (var sklad in resultTDLes)
                    {
                        worksheet2.Cells[row, 1].Value = sklad.WarehouseOwner;
                        worksheet2.Cells[row, 2].Value = sklad.NumTDLes;
                        worksheet2.Cells[row, 3].Value= sklad.DocType;
                        worksheet2.Cells[row, 4].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet2.Cells[row, 4].Value = sklad.DocumentDate;
                        worksheet2.Cells[row, 5].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet2.Cells[row, 5].Value = sklad.ServerProcessingDateTime;
                        worksheet2.Cells[row, 6].Value = sklad.Percon;
                        worksheet2.Cells[row, 7].Value = sklad.DaysDifference;
                        row++;
                    }

                    for (int i = 2; i <= row; i++)
                    {
                        for (int column = 1; column <= 7; column++)
                        {
                            ExcelRange cell = worksheet2.Cells[i, column];
                            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    package.Save();
                }
            }
        }
    }
}
