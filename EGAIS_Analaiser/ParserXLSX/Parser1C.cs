using EGAIS_Analaiser.Data;
using EGAIS_Analaiser.Model.IC;
using OfficeOpenXml;

namespace EGAIS_Analaiser.ParserXLSX
{
    public static class Parser1C
    {
        public static void FullShort1C(string filePath) // парсинг файла-выгрузки из 1с обобщенной информации
        {
            using var package = new ExcelPackage(new FileInfo(filePath));

            var worksheet = package.Workbook.Worksheets.First();

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            using var dbContext = new UserContext();

            for (int row = 5; row <= rowCount; row++)
            {
                Dictionary<string, string> subdivisionReplacements = new Dictionary<string, string>
                        {
                            { "\"Лесопункт\"", "Лесопункт Дятловский лесхоз" },
                            { "ДОЦ Вензовец", "ДОЦ Вензовец Дятловский лесхоз" },
                            { "Дятловский лесхоз", "Станция отгрузки Дятловский лесхоз" }
                        };

                var subdivision = worksheet.Cells[row, 1].Value?.ToString();

                if (subdivisionReplacements.ContainsKey(subdivision))
                {
                    subdivision = subdivisionReplacements[subdivision];
                }

                var record = new FullShort1C
                {
                    Subdivision = subdivision,
                    Procurement = (decimal.TryParse(worksheet.Cells[row, 3].Value?.ToString(), out decimal procurement) ? procurement : 0) * 1000,
                    SelfConsumption = (decimal.TryParse(worksheet.Cells[row, 7].Value?.ToString(), out decimal selfConsumption) ? selfConsumption : 0) * 1000,
                    Sale = (decimal.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out decimal sale) ? sale : 0) * 1000,
                    Processing = (decimal.TryParse(worksheet.Cells[row, 9].Value?.ToString(), out decimal processing) ? processing : 0) * 1000,
                    Balance = (decimal.TryParse(worksheet.Cells[row, 12].Value?.ToString(), out decimal balance) ? balance : 0) * 1000
                };

                dbContext.FullShort1Cs.Add(record);
            }

            dbContext.SaveChanges();
        }

        public static void Parser1c(string filePath) // метод для сбора данных остатков из разных файлов - не используется в текщий момент
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));

            var worksheet = package.Workbook.Worksheets.First();

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            using var dbContext = new UserContext();

            for (int row = 8; row <= rowCount - 6; row++)
            {
                var record = new Remains1C
                {
                    WarehouseOwner = Path.GetFileNameWithoutExtension(filePath),
                    Product = worksheet.Cells[row, 2].Value?.ToString(),
                    Remainder = decimal.TryParse(worksheet.Cells[row, 12].Value?.ToString(), out decimal remainder) ? remainder : 0m
                };

                dbContext.Remains1Cs.Add(record);
            }

            dbContext.SaveChanges();
        } 

        public static void ParserRemainsDOC(string filePath) // метод для сбора данных остатков структурных подраздеений для которых требуется переформатирование исходного файла данных
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));

            var worksheet = package.Workbook.Worksheets.First();

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            using var dbContext = new UserContext();

            for (int row = 1; row <= rowCount; row++)
            {
                var record = new Remains1C
                {
                    WarehouseOwner = Path.GetFileNameWithoutExtension(filePath),
                    Product = worksheet.Cells[row, 4].Value?.ToString(),
                    Remainder = decimal.TryParse(worksheet.Cells[row, 14].Value?.ToString(), out decimal remainder) ? remainder : 0m
                };

                dbContext.Remains1Cs.Add(record);
            }

            dbContext.SaveChanges();
        }
    }
}
