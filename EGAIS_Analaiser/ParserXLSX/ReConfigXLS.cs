using OfficeOpenXml;

namespace EGAIS_Analaiser.ParserXLSX
{
    public class ReConfigXLS
    {
        public static void ConvertXLS(string FileDocPath)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FileDocPath))) // обрабатываем файл с остатками по ДОЦ 1С
            {
                List<string> keywords = new List<string> { "лесомат", "дрова", "балансы", "техсырье", "фенер" };

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = rowCount; row >= 1; row--)
                {
                    // Проверка, пуст ли второй столбец в строке
                    if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 3].Value?.ToString()))
                    {
                        worksheet.DeleteRow(row);
                        continue;
                    }

                    // Получаем содержимое ячейки в нижнем регистре
                    string cellValue = worksheet.Cells[row, 4].Value?.ToString().ToLower();

                    // Проверка, содержит ли ячейка ключевые слова
                    bool containsKeyword = cellValue != null && keywords.Any(keyword => cellValue.Contains(keyword));

                    // Если ячейка не содержит ключевых слов, удаляем строку
                    if (!containsKeyword)
                    {
                        worksheet.DeleteRow(row);
                    }
                }

                // Сброс стиля ячеек
                worksheet.Cells.Style.Font.Bold = false;
                worksheet.Cells.Style.Font.Italic = false;
                worksheet.Cells.Style.Font.Strike = false;

                // Сохранить изменения
                package.Save();
            }
        }

    }
}
