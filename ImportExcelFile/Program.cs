using System;
using System.IO;
using OfficeOpenXml;
using System.Data;
using Microsoft.Extensions.Configuration;

namespace ImportExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = Directory.GetCurrentDirectory();
            IConfiguration config = new ConfigurationBuilder()
                        .SetBasePath(path)
                        .AddJsonFile("appsettings.json", true, true)
                        .Build();

            var fileName = config.GetSection("AppSettings:fileName").Value;
            var filePath = config.GetSection("AppSettings:filePath").Value;

            var data = ImportExcelFile($"{filePath}\\{fileName}");

            Console.WriteLine("Your data has been imported!");
            Console.ReadKey();
        }

        /// <summary>
        /// Import data from a excel file
        /// </summary>
        /// <param name="path">Execel file path</param>
        /// <returns>DataTable with excel file data</returns>
        private static DataTable ImportExcelFile(string path)
        {
            using (var file = File.OpenRead(path))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    using (var package = new ExcelPackage(ms))
                    {
                        var sheet = package.Workbook.Worksheets[1];
                        var dt = GetDataTable(package);

                        var rowsNumbler = sheet.Dimension?.Rows;
                        var columnsNumbler = sheet.Dimension?.Columns;
                        for (int row = 2; row <= rowsNumbler; row++)
                        {
                            var newRow = dt.NewRow();
                            for (int col = 1; col <= columnsNumbler; col++)
                            {
                                newRow[col - 1] = sheet.Cells[row, col].Value;
                            }
                            dt.Rows.Add(newRow);
                        }
                        return dt;
                    }
                }
            }
        }

        /// <summary>
        /// It creates a DataTable with the same colomns present in the excel file.
        /// </summary>
        /// <param name="excelPackage">Excel package file</param>
        /// <returns></returns>
        private static DataTable GetDataTable(ExcelPackage excelPackage)
        {
            int index = 1;
            var sheet = excelPackage.Workbook.Worksheets[index];
            int? columnsNumber = sheet.Dimension?.Columns;
            var dt = new DataTable(sheet.Name);
            for (int i = 1; i <= columnsNumber; i++)
            {
                dt.Columns.Add(sheet.Cells[index, i].Value.ToString());
            }

            return dt;
        }
    }
}
