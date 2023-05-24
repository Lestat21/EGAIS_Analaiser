using EGAIS_Analaiser.Data;
using EGAIS_Analaiser.Model.Egais;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EGAIS_Analaiser.View
{
    public class SaveToFile
    {
        public static void ToFile(string newFilePath)
        {
            using var dbContext = new UserContext();

            #region == Linq выборка данных для анализа

            var _resultsEGAIS = dbContext.Remains
                     .GroupBy(record => record.WarehouseOwner)
                     .Select(group => new
                     {
                         WarehouseOwner = group.Key,
                         TotalVolume = group.Sum(record => record.Volume)
                     })
                     .OrderBy(result => result.WarehouseOwner);

            var _zagotovkaEGAIS = dbContext.Zagotovkas
                     .GroupBy(record => record.Forestry)
                     .Select(group => new
                     {
                         WarehouseOwner = group.Key,
                         TotalVolume = group.Sum(record => record.OperationalAccountingTotal)
                     })
                     .OrderBy(result => result.WarehouseOwner);

            var _selling = dbContext.Sellings
                    .GroupBy(record => record.Division)
                    .Select(group => new
                    {
                        Division = group.Key,
                        TotalVolume = group.Sum(record => record.Volume)
                    })
                    .OrderBy(result => result.Division);

            var _result1C = dbContext.FullShort1Cs
                      .Select(record => new
                      {
                          Subdivision = record.Subdivision,
                          Zagotovka = record.Procurement,
                          Balance = record.Balance,
                          Sale = record.Sale,
                          SelfConsumption = record.SelfConsumption,
                          Processing = record.Processing
                      }).OrderBy(result => result.Subdivision).ToList();

            var _resultTDLes = dbContext.TDLes
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

            var _skladEGAIS = dbContext.Sklads
                   .Select(record => new
                   {
                       Sklad = record.Name,
                       WarehouseOwner = record.WarehouseOwner,
                       CreateData = record.OpenDate,
                       WarehouseType = record.WarehouseType,
                       CreateUser = record.CreateUser
                   }).ToList();

            var _remainsBySlad = dbContext.Remains
                    .GroupBy(record => record.Warehouse)
                    .Select(group => new
                    {
                        Sklad = group.Key,
                        TotalVolume = group.Sum(record => record.Volume)
                    }).ToList();

            var _skladToRemove = _skladEGAIS
                   .Where(s => !_remainsBySlad.Any(r => r.Sklad == s.Sklad))
                   .OrderBy(s => s.WarehouseOwner)
                   .ToList();


            DateTime targetDate = new DateTime(2023, 1, 1);

            var _countSlklad = _skladToRemove
                    .Where(group => group.CreateData < targetDate)
                    .GroupBy(item => item.WarehouseOwner)
                    .Select(group => new
                    {
                        Name = group.Key,
                        Count = group.Count()
                    });

            var _tmpSklad = from sklad in _skladEGAIS
                            join remainsBySlad in _remainsBySlad on sklad.Sklad equals remainsBySlad.Sklad into skladGroup
                            from remainsBySlad in skladGroup.DefaultIfEmpty()
                            where sklad.WarehouseType.Contains("Временный")
                            select new
                            {
                                WarehouseOwner = sklad.WarehouseOwner,
                                Warehouse = sklad.Sklad,
                                WarehouseType = sklad.WarehouseType,
                                Volume = remainsBySlad?.TotalVolume,
                                CreateUser = sklad.CreateUser,
                                CreateData = sklad.CreateData
                            };

            #endregion

            using var package = new ExcelPackage(new FileInfo(newFilePath));

            #region == Первый лист ОБЩАЯ АНАЛИТИКА

            var worksheet = package.Workbook.Worksheets.Add("ОБЩАЯ АНАЛИТИКА");

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

            foreach (var result in _resultsEGAIS)
            {
                var r1c = _result1C.FirstOrDefault(p => p.Subdivision == result.WarehouseOwner)?.Balance ?? 0;

                worksheet.Cells[row, 1].Value = result.WarehouseOwner;
                worksheet.Cells[row, 2].Value = result.TotalVolume;
                worksheet.Cells[row, 3].Value = r1c;
                worksheet.Cells[row, 4].Value = result.TotalVolume / r1c * 100;
                worksheet.Cells[row, 5].Value = r1c - result.TotalVolume;
                row++;
            }

            row = 3;
            foreach (var result in _zagotovkaEGAIS)
            {
                var r1c = _result1C.FirstOrDefault(p => p.Subdivision == result.WarehouseOwner)?.Zagotovka ?? 0;

                worksheet.Cells[row, 6].Value = result.TotalVolume;
                worksheet.Cells[row, 7].Value = r1c;
                worksheet.Cells[row, 8].Value = result.TotalVolume / r1c * 100;
                worksheet.Cells[row, 9].Value = r1c - result.TotalVolume;
                row++;
            }

            row = 3;
            foreach (var result in _selling)
            {
                var r1c = (_result1C.FirstOrDefault(p => p.Subdivision == result.Division)?.Sale ?? 0) +
                    (_result1C.FirstOrDefault(p => p.Subdivision == result.Division)?.SelfConsumption ?? 0) +
                    (_result1C.FirstOrDefault(p => p.Subdivision == result.Division)?.Processing ?? 0);

                worksheet.Cells[row, 10].Value = result.TotalVolume;
                worksheet.Cells[row, 11].Value = r1c;
                worksheet.Cells[row, 12].Value = result.TotalVolume / r1c * 100;
                worksheet.Cells[row, 13].Value = r1c - result.TotalVolume;
                row++;
            }

            worksheet.Cells[row, 1].Value = "";
            worksheet.Cells[row, 2].Value = _resultsEGAIS.Sum(r => r.TotalVolume);
            worksheet.Cells[row, 3].Value = _result1C.Sum(r => r.Balance);
            worksheet.Cells[row, 4].Value = _resultsEGAIS.Sum(r => r.TotalVolume) / _result1C.Sum(r => r.Balance) * 100;
            worksheet.Cells[row, 5].Value = _result1C.Sum(r => r.Balance) - _resultsEGAIS.Sum(r => r.TotalVolume);
            worksheet.Cells[row, 6].Value = _zagotovkaEGAIS.Sum(r => r.TotalVolume);
            worksheet.Cells[row, 7].Value = _result1C.Sum(r => r.Zagotovka);
            worksheet.Cells[row, 8].Value = _zagotovkaEGAIS.Sum(r => r.TotalVolume) / _result1C.Sum(r => r.Zagotovka) * 100;
            worksheet.Cells[row, 9].Value = _result1C.Sum(r => r.Zagotovka) - _zagotovkaEGAIS.Sum(r => r.TotalVolume);
            worksheet.Cells[row, 10].Value = _selling.Sum(r => r.TotalVolume);
            worksheet.Cells[row, 11].Value = _result1C.Sum(r => r.SelfConsumption) + _result1C.Sum(r => r.Sale) + _result1C.Sum(r => r.Processing);
            worksheet.Cells[row, 12].Value = _selling.Sum(r => r.TotalVolume) / (_result1C.Sum(r => r.SelfConsumption) + _result1C.Sum(r => r.Sale) + _result1C.Sum(r => r.Processing)) * 100;
            worksheet.Cells[row, 13].Value = _result1C.Sum(r => r.SelfConsumption) + _result1C.Sum(r => r.Sale) + _result1C.Sum(r => r.Processing) - _selling.Sum(r => r.TotalVolume);
            worksheet.Cells[row, 14].Value = Math.Abs((_resultsEGAIS.Sum(r => r.TotalVolume) / _result1C.Sum(r => r.Balance) * 100) - 1) +
                Math.Abs((_zagotovkaEGAIS.Sum(r => r.TotalVolume) / (_result1C.Sum(r => r.Zagotovka) * 100) - 1)) +
                Math.Abs((_selling.Sum(r => r.TotalVolume) / (_result1C.Sum(r => r.SelfConsumption) + _result1C.Sum(r => r.Sale) + _result1C.Sum(r => r.Processing) * 100) - 1));
            #endregion

            Console.WriteLine($"{DateTime.Now} - Общая аналитика создана и записана в файл");

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

            #region == Пустые склады ==
            var worksheet1 = package.Workbook.Worksheets.Add("ПУСТЫЕ СКЛАДЫ");

            worksheet1.Column(1).Width = 30;
            worksheet1.Column(2).Width = 70;
            worksheet1.Column(3).Width = 15;
            worksheet1.Column(5).Width = 30;
            worksheet1.Column(6).Width = 15;

            worksheet1.Cells[2, 1].Value = "Лесничество";
            worksheet1.Cells[2, 2].Value = "Наименование склада";
            worksheet1.Cells[2, 3].Value = "Дата открытия";

            worksheet1.Cells[2, 5].Value = "Лесничество";
            worksheet1.Cells[2, 6].Value = "Количество";


            row = 3;
            foreach (var sklad in _skladToRemove)
            {
                worksheet1.Cells[row, 1].Value = sklad.WarehouseOwner;
                worksheet1.Cells[row, 2].Value = sklad.Sklad;
                worksheet1.Cells[row, 3].Style.Numberformat.Format = "dd.MM.yyyy";
                worksheet1.Cells[row, 3].Value = sklad.CreateData;
                row++;
            }


            for (int i = 2; i < row; i++)
            {
                for (int column = 1; column <= 3; column++)
                {
                    ExcelRange cell = worksheet1.Cells[i, column];
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
            }

            row = 3;
            foreach(var sklad in _countSlklad)
            {
                worksheet1.Cells[row, 5].Value = sklad.Name;
                worksheet1.Cells[row, 6].Value = sklad.Count;
                row++;
            }

            for (int i = 2; i < row; i++)
            {
                for (int column = 5; column <= 6; column++)
                {
                    ExcelRange cell = worksheet1.Cells[i, column];
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
            }

            #endregion

            Console.WriteLine($"{DateTime.Now} - Выбраны пустые склады и записаны в файл");

            #region == НАРУШЕНИЯ УЧЕТА

            var worksheet2 = package.Workbook.Worksheets.Add("Нарушения. Учет ТДЛЕС");

            worksheet2.Column(1).Width = 30;
            worksheet2.Column(2).Width = 30;
            worksheet2.Column(3).Width = 40;
            worksheet2.Column(4).Width = 15;
            worksheet2.Column(5).Width = 15;
            worksheet2.Column(6).Width = 20;
            worksheet2.Column(7).Width = 15;
            worksheet2.Cells[2, 1].Value = "Лесничество";
            worksheet2.Cells[2, 2].Value = "Номер документа";
            worksheet2.Cells[2, 3].Value = "Тип документа";
            worksheet2.Cells[2, 4].Value = "Дата документа";
            worksheet2.Cells[2, 5].Value = "Время сервера";
            worksheet2.Cells[2, 6].Value = "Пользователь";
            worksheet2.Cells[2, 7].Value = "Расхождение";

            row = 3;
            foreach (var sklad in _resultTDLes)
            {
                worksheet2.Cells[row, 1].Value = sklad.WarehouseOwner;
                worksheet2.Cells[row, 2].Value = sklad.NumTDLes;
                worksheet2.Cells[row, 3].Value = sklad.DocType;
                worksheet2.Cells[row, 4].Style.Numberformat.Format = "dd.MM.yyyy";
                worksheet2.Cells[row, 4].Value = sklad.DocumentDate;
                worksheet2.Cells[row, 5].Style.Numberformat.Format = "dd.MM.yyyy";
                worksheet2.Cells[row, 5].Value = sklad.ServerProcessingDateTime;
                worksheet2.Cells[row, 6].Value = sklad.Percon;
                worksheet2.Cells[row, 7].Value = sklad.DaysDifference;
                row++;
            }

            for (int i = 2; i < row; i++)
            {
                for (int column = 1; column <= 7; column++)
                {
                    ExcelRange cell = worksheet2.Cells[i, column];
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
            }
            #endregion

            Console.WriteLine($"{DateTime.Now} - Выбраны сотрудники,допустившие нарушения учета и записаны в файл");

            #region == Временные склады == 

            var worksheet3 = package.Workbook.Worksheets.Add("ВРЕМЕННЫЕ СКЛАДЫ");

            worksheet3.Column(1).Width = 30;
            worksheet3.Column(2).Width = 40;
            worksheet3.Column(3).Width = 35;
            worksheet3.Column(4).Width = 10;
            worksheet3.Column(5).Width = 20;
            worksheet3.Column(6).Width = 15;

            worksheet3.Cells[2, 1].Value = "Лесничество";
            worksheet3.Cells[2, 2].Value = "Наименование склада";
            worksheet3.Cells[2, 3].Value = "Тип склада";
            worksheet3.Cells[2, 4].Value = "Объем остатков";
            worksheet3.Cells[2, 5].Value = "Кто создал";
            worksheet3.Cells[2, 6].Value = "Дата создания";

            row = 3;

            foreach (var sklad in _tmpSklad)
            {
                worksheet3.Cells[row, 1].Value = sklad.WarehouseOwner;
                worksheet3.Cells[row, 2].Value = sklad.Warehouse;
                worksheet3.Cells[row, 3].Value = sklad.WarehouseType;
                worksheet3.Cells[row, 4].Value = sklad.Volume;
                worksheet3.Cells[row, 5].Value = sklad.CreateUser;
                worksheet3.Cells[row, 6].Style.Numberformat.Format = "dd.MM.yyyy";
                worksheet3.Cells[row, 6].Value = sklad.CreateData;
                row++;
            }

            for (int i = 2; i < row; i++)
            {
                for (int column = 1; column <= 6; column++)
                {
                    ExcelRange cell = worksheet3.Cells[i, column];
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
            }








            #endregion

            Console.WriteLine($"{DateTime.Now} - Выбраны временные склады и записаны в файл");

            package.Save();
        }
    }
}
