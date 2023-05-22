using EGAIS_Analaiser.Data;
using EGAIS_Analaiser.Model.Egais;
using OfficeOpenXml;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace EGAIS_Analaiser.ParserXLSX
{
    public static class ParserEGAIS
    {
        public static void ParserEgais(string filePath) // парсер остатков ЕГАИС
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.First();

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                using (var dbContext = new UserContext())
                {
                    for (int row = 2; row <= rowCount - 1; row++)
                    {
                        var record = new Remains
                        {
                            LesHoz = worksheet.Cells[row, 2].Value?.ToString(),
                            WarehouseOwner = worksheet.Cells[row, 1].Value?.ToString(),
                            MinistryLevel = worksheet.Cells[row, 3].Value?.ToString(),
                            GPLHO_Level = worksheet.Cells[row, 4].Value?.ToString(),
                            Warehouse = worksheet.Cells[row, 5].Value?.ToString(),
                            ForestQuartalNumber = worksheet.Cells[row, 6].Value?.ToString(),
                            TaxDivisionNumber = worksheet.Cells[row, 7].Value?.ToString(),
                            Nomenclature = worksheet.Cells[row, 8].Value?.ToString(),
                            VolumeDeterminationMethod = worksheet.Cells[row, 9].Value?.ToString(),
                            Quantity = int.TryParse(worksheet.Cells[row, 10].Value?.ToString(), out int quantity) ? quantity : 0,
                            Volume = decimal.TryParse(worksheet.Cells[row, 11].Value?.ToString().Replace(".", ","), out decimal volume) ? volume : 0,
                            TreeSpecies = worksheet.Cells[row, 12].Value?.ToString(),
                            Assortment = worksheet.Cells[row, 13].Value?.ToString(),
                            DiameterGroup = worksheet.Cells[row, 14].Value?.ToString(),
                            Diameter = worksheet.Cells[row, 15].Value?.ToString(),
                            Length = worksheet.Cells[row, 16].Value?.ToString()
                        };

                        dbContext.Remains.Add(record);
                    }

                    dbContext.SaveChanges();
                }
            }
        }

        public static void ParserSellingEGAIS(string filePath) //парсер реализации ЕГАИС
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                using (var dbContext = new UserContext())
                {
                    for (int row = 2; row <= rowCount - 1; row++)
                    {

                        Dictionary<string, string> subdivisionReplacements = new Dictionary<string, string>
                        {
                            { "Козловщинское", "Козловщинское лесничество" },
                            { "Охоновское", "Охоновское лесничество" },
                            { "Вензовецкое", "Вензовецкое лесничество" },
                            { "Новоельнянское", "Новоельнянское лесничество" },
                            { "Роготновское", "Роготновское лесничество" },
                            { "Демьяновичское", "Демьяновичское лесничество" },
                            { "Гезгаловское", "Гезгаловское лесничество" },
                            { "Леоновичское", "Леоновичское лесничество" },
                            { "Руда-Яворское", "Руда-Яворское лесничество" },
                            { "Лесопункт Дятловский лесхоз", "Лесопункт Дятловский лесхоз" },
                            { "ДОЦ Вензовец", "ДОЦ Вензовец Дятловский лесхоз" },
                            { "Станция отгрузки", "Станция отгрузки Дятловский лесхоз" }
                        };

                        var subdivision = worksheet.Cells[row, 1].Value.ToString();

                        if (subdivisionReplacements.ContainsKey(subdivision))
                        {
                            subdivision = subdivisionReplacements[subdivision];
                        }

                        Selling selling = new Selling();

                        selling.Division = subdivision; // Структурное подразделение
                        selling.WarehouseOwner = worksheet.Cells[row, 5].Value.ToString(); // Владелец склада
                        selling.ForestQuartNumber = worksheet.Cells[row, 6].Value.ToString(); // Номер лесного квартала
                        selling.TaxationAreaNumber = worksheet.Cells[row, 7].Value.ToString(); // Номер таксационного выдела
                        selling.OperationWarehouse = worksheet.Cells[row, 8].Value.ToString(); // Склад операции
                        selling.Nomenclature = worksheet.Cells[row, 9].Value.ToString(); // Номенклатура
                        selling.Quantity = decimal.Parse(worksheet.Cells[row, 10].Value.ToString()); // Кол-во
                        selling.Volume = decimal.Parse(worksheet.Cells[row, 11].Value.ToString().Substring(1).Replace(".", ",")); // Объем
                        selling.DocumentType = worksheet.Cells[row, 12].Value.ToString(); // Тип документа
                        selling.DocumentNumber = worksheet.Cells[row, 13].Value.ToString(); // Номер документа
                        selling.DocumentDate = DateTime.Parse(worksheet.Cells[row, 14].Value.ToString()); // Дата документа
                        selling.Shipper = worksheet.Cells[row, 15].Value.ToString(); // Грузоотправитель
                        selling.Consignee = worksheet.Cells[row, 16].Value.ToString(); // Грузополучатель
                        selling.CounterpartyWarehouse = worksheet.Cells[row, 17].Value?.ToString(); // Склад контрагента
                        selling.Employee = worksheet.Cells[row, 18].Value.ToString(); // Сотрудник
                        selling.Reason = worksheet.Cells[row, 19].Value.ToString(); // Основание
                        selling.ReasonNumber = worksheet.Cells[row, 20].Value.ToString(); // Номер основания
                        selling.ReasonDate = DateTime.Parse(worksheet.Cells[row, 21].Value.ToString()); // Дата основания
                        selling.Status = worksheet.Cells[row, 22].Value.ToString(); // Статус
                        selling.ServerProcessingDateTime = DateTime.Parse(worksheet.Cells[row, 23].Value.ToString()); // Дата и время обработки на сервере
                        selling.CreationUser = worksheet.Cells[row, 24].Value.ToString(); // Пользователь создания
                        selling.Transport = worksheet.Cells[row, 27].Value.ToString(); // Транспорт
                        selling.Trailer = worksheet.Cells[row, 28].Value.ToString(); // Прицеп
                        selling.VolumeDeterminationMethod = worksheet.Cells[row, 29].Value.ToString(); // Метод определения объема

                        dbContext.Sellings.Add(selling);
                    }
                    dbContext.SaveChanges();
                }
            }
        }

        public static void ParserZagotovkaEGAIS(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.First();

                int rowCount = worksheet.Dimension.Rows;

                using (var dbContext = new UserContext())
                {
                    for (int row = 2; row <= rowCount - 1; row++)
                    {
                        Dictionary<string, string> subdivisionReplacements = new Dictionary<string, string>
                        {
                            { "Козловщинское", "Козловщинское лесничество" },
                            { "Охоновское", "Охоновское лесничество" },
                            { "Вензовецкое", "Вензовецкое лесничество" },
                            { "Новоельнянское", "Новоельнянское лесничество" },
                            { "Роготновское", "Роготновское лесничество" },
                            { "Демьяновичское", "Демьяновичское лесничество" },
                            { "Гезгаловское", "Гезгаловское лесничество" },
                            { "Леоновичское", "Леоновичское лесничество" },
                            { "Руда-Яворское", "Руда-Яворское лесничество" },
                            { "Лесопункт Дятловский лесхоз", "Лесопункт Дятловский лесхоз" },
                            { "ДОЦ Вензовец", "ДОЦ Вензовец Дятловский лесхоз" },
                            { "Станция отгрузки", "Станция отгрузки Дятловский лесхоз" }
                        };

                        var subdivision = worksheet.Cells[row, 20].Value?.ToString();

                        if (subdivisionReplacements.ContainsKey(subdivision))
                        {
                            subdivision = subdivisionReplacements[subdivision];
                        }

                        var record = new Zagotovka
                        {
                            OperationalAccountingTotal = decimal.TryParse(worksheet.Cells[row, 11].Value?.ToString().Replace(".", ","), out decimal operationalAccountingTotal) ? operationalAccountingTotal : 0m,
                            LoggingMethod = worksheet.Cells[row, 12].Value?.ToString(),
                            Forestry = subdivision,
                        };

                        dbContext.Zagotovkas.Add(record);
                    }

                    dbContext.SaveChanges();
                }
            }
        }

        public static void ParserSkladEGAIS(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.First();

                int rowCount = worksheet.Dimension.Rows;

                using (var dbContext = new UserContext())
                {
                    for (int row = 2; row <= rowCount - 1; row++)
                    {
                        var record = new Sklad();
                        record.Name = worksheet.Cells[row, 1].Value?.ToString();
                        record.WarehouseOwner = worksheet.Cells[row, 2].Value?.ToString();
                        record.ForestQuarterNumber = worksheet.Cells[row, 6].Value?.ToString();
                        record.TaxationPlotNumber = worksheet.Cells[row, 7].Value?.ToString();
                        record.WarehouseAddress = worksheet.Cells[row, 8].Value?.ToString();
                        record.WarehouseType = worksheet.Cells[row, 9].Value?.ToString();
                        record.ActivityType = worksheet.Cells[row, 10].Value?.ToString();
                        record.LoggingSite = worksheet.Cells[row, 11].Value?.ToString();
                        record.Latitude = double.TryParse(worksheet.Cells[row, 12].Value?.ToString(), out double latitude) ? latitude : 0.0;
                        record.Longitude = double.TryParse(worksheet.Cells[row, 13].Value?.ToString(), out double longitude) ? longitude : 0.0;
                        record.OpenDate = DateTime.Parse(worksheet.Cells[row, 14].Value?.ToString());
                        record.CloseDate = DateTime.TryParse(worksheet.Cells[row, 15].Value?.ToString(), out DateTime parsedDate) ? parsedDate : (DateTime?)null;
                        record.Status = worksheet.Cells[row, 16].Value?.ToString();
                        record.SkladId = int.TryParse(worksheet.Cells[row, 17].Value?.ToString(), out int id) ? id : 0;
                        record.CreateDate = DateTime.Parse(worksheet.Cells[row, 18].Value?.ToString());
                        record.CreateUser = worksheet.Cells[row, 19].Value?.ToString();
                        record.ModifyDate = DateTime.TryParse(worksheet.Cells[row, 20].Value?.ToString(), out DateTime parsedDate1) ? parsedDate1 : (DateTime?)null;
                        record.ModifyUser = worksheet.Cells[row, 21].Value?.ToString();
                        record.WoodHarvestedInRadioactiveContaminationZone = bool.TryParse(worksheet.Cells[row, 22].Value?.ToString(), out bool woodHarvestedInRadioactiveContaminationZone) ? woodHarvestedInRadioactiveContaminationZone : false;
                        record.StatusCode = worksheet.Cells[row, 23].Value?.ToString();

                        dbContext.Sklads.Add(record);
                    }

                    dbContext.SaveChanges();
                }
            }
        }

        public static void ParserTDLes(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.First();

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                using (var dbContext = new UserContext())
                {
                    for (int row = 2; row <= rowCount - 1; row++)
                    {
                        var record = new TDLes();

                        record.WarehouseOwner = worksheet.Cells[row, 1].Value?.ToString();
                        record.DocumentType = worksheet.Cells[row, 5].Value?.ToString();
                        record.DocumentNumber = worksheet.Cells[row, 6].Value?.ToString();
                        record.RelatedDocumentNumber = worksheet.Cells[row, 7].Value?.ToString();
                        record.DocumentDate = Convert.ToDateTime(worksheet.Cells[row, 8].Value);
                        record.Employee = worksheet.Cells[row, 9].Value?.ToString();
                        record.OperationWarehouse = worksheet.Cells[row, 10].Value?.ToString();
                        record.ForestQuarterNumber = worksheet.Cells[row, 11].Value?.ToString();
                        record.TaxationPlotNumber = worksheet.Cells[row, 12].Value?.ToString();
                        record.CounterpartyWarehouse = worksheet.Cells[row, 13].Value?.ToString();
                        record.BasisDocument = worksheet.Cells[row, 14].Value?.ToString();
                        record.BasisDocumentNumber = worksheet.Cells[row, 15].Value?.ToString();
                        record.BasisDocumentDate = DateTime.TryParse(worksheet.Cells[row, 16].Value?.ToString(), out DateTime parsedDate0) ? parsedDate0 : (DateTime?)null;
                        record.Status = worksheet.Cells[row, 17].Value?.ToString();
                        record.Shipper = worksheet.Cells[row, 18].Value?.ToString();
                        record.Carrier = worksheet.Cells[row, 19].Value?.ToString();
                        record.Consignee = worksheet.Cells[row, 20].Value?.ToString();
                        record.ConsigneeFlag = worksheet.Cells[row, 21].Value?.ToString();
                        record.CreatedByUser = worksheet.Cells[row, 22].Value?.ToString();
                        record.ModifiedByUser = worksheet.Cells[row, 23].Value?.ToString();
                        record.ServerProcessingDateTime = Convert.ToDateTime(worksheet.Cells[row, 24].Value);
                        record.ModifiedDate = DateTime.TryParse(worksheet.Cells[row, 25].Value?.ToString(), out DateTime parsedDate1) ? parsedDate1 : (DateTime?)null;
                        record.ChangeBasis = worksheet.Cells[row, 26].Value?.ToString();
                        record.AdjustmentBasisDocumentNumber = worksheet.Cells[row, 27].Value?.ToString();
                        record.AdjustmentBasisDocumentDate = DateTime.TryParse(worksheet.Cells[row, 28].Value?.ToString(), out DateTime parsedDate) ? parsedDate : (DateTime?)null; 

                        dbContext.TDLes.Add(record);
                    }

                    dbContext.SaveChanges();
                }
            }
        }
    }
}
