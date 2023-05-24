using OfficeOpenXml;

namespace EGAIS_Analaiser.ParserXLSX
{
    public class ReConfigXLS
    {
        public static void ConvertXLS(string FileDocPath) // обрабатываем файл с остатками по ДОЦ 1С/Станция погрузки
        {
            using ExcelPackage package = new ExcelPackage(new FileInfo(FileDocPath));

            List<string> keywords = new List<string> { "лесомат", "дрова", "балансы", "техсырье", "фенер" };

            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = rowCount; row >= 1; row--)
            {
                if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 3].Value?.ToString()))
                {
                    worksheet.DeleteRow(row);
                    continue;
                }

                string? cellValue = worksheet.Cells[row, 4].Value?.ToString().ToLower();

                bool containsKeyword = cellValue != null && keywords.Any(keyword => cellValue.Contains(keyword));

                if (!containsKeyword)
                {
                    worksheet.DeleteRow(row);
                }
            }

            worksheet.Cells.Style.Font.Bold = false;
            worksheet.Cells.Style.Font.Italic = false;
            worksheet.Cells.Style.Font.Strike = false;

            package.Save();
        }
    }
}
