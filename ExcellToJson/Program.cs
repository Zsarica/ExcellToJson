using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcellToJson
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = ConfigurationManager.AppSettings["excelFilePath"].ToString().Trim();//@"C:\Users\Zekeriya SARICA\Desktop\Il_Ilce.xlsx";
            string jsonFilePath = ConfigurationManager.AppSettings["jsonFilePath"].ToString().Trim();//@"C:\Users\Zekeriya SARICA\Desktop\Output.txt";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Excel dosyasını oku
            var data = ReadExcelFile(excelFilePath);

            // JSON formatına dönüştür
            string json = JsonConvert.SerializeObject(data, Formatting.Indented);

            // JSON'u dosyaya yaz
            File.WriteAllText(jsonFilePath, json);

            Console.WriteLine("JSON dosyası başarıyla oluşturuldu.");
            Console.ReadKey();  
        }

        static List<Dictionary<string, object>> ReadExcelFile(string filePath)
        {
            var result = new List<Dictionary<string, object>>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowDict = new Dictionary<string, object>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        rowDict[headers[col - 1]] = worksheet.Cells[row, col].Text;
                    }
                    result.Add(rowDict);
                }
            }

            return result;
        }
    }
}
