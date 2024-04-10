using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace GP_EncriptionFiles
{
    public  class HashAlgorithmTest
    {

        public   static void  test()
        {
            string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "download_sample_files"); // Change this to your directory
            string sequentialFolderPath = Utilities.CreateDirectoryIfNotExist(directoryPath, "EncryptionFolderSequentialTestAll");
            string excelFilePath = Path.Combine(sequentialFolderPath, "EncryptionDataSequential.xlsx");

           // string excelFilePath = @"C:\Path\To\Save\Excel\File.xlsx"; // Change this to your desired Excel file path

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("HashTimes");
                worksheet.Cell("A1").Value = "FileName";
                worksheet.Cell("B1").Value = "SHA-1 Time (ms)";
                worksheet.Cell("C1").Value = "SHA-256 Time (ms)";
                worksheet.Cell("D1").Value = "MD5 Time (ms)";

                int row = 2;
                foreach (var filePath in Directory.GetFiles(directoryPath))
                {
                    var fileName = Path.GetFileName(filePath);
                    worksheet.Cell(row, 1).Value = fileName;

                    // SHA-1
                    long sha1Time = MeasureHashTime(() => SHA1.Create(), filePath);
                    worksheet.Cell(row, 2).Value = sha1Time;

                    // SHA-256
                    long sha256Time = MeasureHashTime(() => SHA256.Create(), filePath);
                    worksheet.Cell(row, 3).Value = sha256Time;

                    // MD5
                    long md5Time = MeasureHashTime(() => MD5.Create(), filePath);
                    worksheet.Cell(row, 4).Value = md5Time;

                    row++;
                }

                workbook.SaveAs(excelFilePath);
            }

            Console.WriteLine("Process completed. Check the Excel file.");
        }

        static long MeasureHashTime(Func<HashAlgorithm> createAlgorithm, string filePath)
        {
            using (var algorithm = createAlgorithm())
            {
                using (var stream = File.OpenRead(filePath))
                {
                    var stopwatch = new Stopwatch();
                    stopwatch.Start();
                    byte[] hash = algorithm.ComputeHash(stream);
                    stopwatch.Stop();
                    return stopwatch.ElapsedMilliseconds;
                }
            }
        }
    }
}
