using ConsoleTables;
using OfficeOpenXml;
using System.Data.Common;
using System.Linq.Expressions;

namespace University.ScoreCalculator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string marksPath = GetInputFile("marks.xlsx");

            ConsoleTable results = GetResults(marksPath, columns: [2, 4]);

            Console.WriteLine(results);

            double avgMark = AvgMark(results, 2);

            Console.WriteLine($"Средний балл: {Math.Round(avgMark, 3)}");
        }

        public static string GetInputFile(string fileName)
        {
            string inputPath = Path.Combine(
                Directory.GetParent(Directory.GetCurrentDirectory())?.Parent?.Parent.FullName, 
                "Input", 
                fileName);
            
            return inputPath;
        }

        public static ConsoleTable GetResults(string filePath, string sheetName = "Лист1", params int[] columns)
        {
            if (columns.Length > 5 || columns.Length < 1)
            {
                throw new ArgumentException(nameof(columns));
            }

            if (columns.Any(elem => elem < 1 || elem > 5))
            {
                throw new ArgumentException(nameof(columns));
            }

            var table = new ConsoleTable();

            for (int i = 0; i < columns.Length; i++)
            {
                switch (columns[i])
                {
                    case 1:
                        table.AddColumn(new string[] { "Семестр" });
                        break;
                    case 2:
                        table.AddColumn(new string[] { "Дисциплина" });
                        break;
                    case 3:
                        table.AddColumn(new string[] { "Форма контроля" });
                        break;
                    case 4:
                        table.AddColumn(new string[] { "Оценка" });
                        break;
                    case 5:
                        table.AddColumn(new string[] { "Дата" });
                        break;
                    default:
                        break;
                }
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Лист не найден!");
                    }

                    int rowCount = worksheet.Dimension.Rows;
                    
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string[] selectedRow = new string[columns.Length];
                        for (int i = 0; i < selectedRow.Length; i++)
                        {
                            selectedRow[i] = worksheet.Cells[row, columns[i]].Text;
                        }
                        table.AddRow(selectedRow);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }

            return table;
        }

        public static double AvgMark(ConsoleTable table, int markColumn)
        {
            double avgMark = 0;
            int marks = 0;

            foreach (var mark in table.Rows)
            {
                switch (mark[markColumn - 1].ToString().ToLower())
                {
                    case "удовлетворительно":
                        avgMark += 3;
                        marks++;
                        break;
                    case "хорошо":
                        avgMark += 4;
                        marks++;
                        break;
                    case "отлично":
                        avgMark += 5;
                        marks++;
                        break;
                    default:
                        break;
                }
            }

            avgMark /= marks;

            return avgMark;
        }
    }
}
