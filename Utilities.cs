using ClosedXML.Excel;
using Org.BouncyCastle.Crypto.Engines;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Security;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Collections.Generic;
namespace GP_EncriptionFiles
{
    public class Utilities
    {
        //this path have the malware file sample 
        public static string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "download_sample_files");

        public static string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        public static double EncryptAndHashFile(string inputFile, string outputFile, out string hash, Func<HashAlgorithm> createAlgorithm)
        {

            var stopwatch = Stopwatch.StartNew();

            using (var algorithm = createAlgorithm())
            //using (var sha256 = SHA256.Create())
            using (var inputFileStream = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
            {
                // Compute SHA256 hash of the file data
                byte[] hashBytes = algorithm.ComputeHash(inputFileStream);
                hash = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
            }

            stopwatch.Stop();
            return stopwatch.ElapsedMilliseconds;
        }

        public static double EncryptFileChaCha20(string filePath, out string hash)
        {
            var stopwatch = Stopwatch.StartNew();
            // Read the file
            byte[] fileData = File.ReadAllBytes(filePath);

            // Generate Key and IV
            SecureRandom random = new SecureRandom();
            byte[] key = new byte[32]; // 256 bits
            byte[] iv = new byte[12]; // 96 bits (12 bytes)
            random.NextBytes(key);
            random.NextBytes(iv);

            // Create the ChaCha20 cipher
            IBufferedCipher cipher = CipherUtilities.GetCipher("ChaCha20");
            cipher.Init(true, new ParametersWithIV(new KeyParameter(key), iv));

            // Encrypt the data
            byte[] encryptedData = cipher.DoFinal(fileData);

            hash = BitConverter.ToString(encryptedData).Replace("-", "").ToLowerInvariant();
            // Save or use the encrypted data
            // File.WriteAllBytes("path/to/encrypted/file", encryptedData);

            Console.WriteLine("Encryption complete.");
            stopwatch.Stop();
            return stopwatch.Elapsed.TotalSeconds;
        }
        public static async Task<List<(string FileName, double EncryptionTime, string Hash, long FileSize)>> EncryptAndHashDirectoryAsync(string directoryPath, Func<HashAlgorithm> createAlgorithm)
        {
            var results = new List<(string FileName, double EncryptionTime, string Hash, long FileSize)>();
            var files = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                string encryptedFile = $"{file}.enc";

                var result = await Task.Run(() =>
                {
                    string hash;
                    double duration = EncryptAndHashFile(file, encryptedFile, out hash, createAlgorithm);
                    long fileSize = new FileInfo(file).Length;
                    return (FileName: file, EncryptionTime: duration, Hash: hash, FileSize: fileSize);
                });

                results.Add(result);
            }

            return results;
        }

        public static void WriteToExcel(List<(string FileName, double EncryptionTime, string Hash, long FileSize)> fileData, string outputPath, string algorithm, Boolean isSequintal)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Encryption Data");
                worksheet.Cell(1, 1).Value = "File Name";
                worksheet.Cell(1, 4).Value = $"{(isSequintal ? "Sequintal" : "Paralllel")}{algorithm}Encryption Time (Milliseconds)";
                worksheet.Cell(1, 3).Value = $"{(isSequintal ? "Sequintal" : "Paralllel")}Hash {algorithm}";
                worksheet.Cell(1, 2).Value = "File Size (Bytes)";
                int row = 2;
                foreach (var data in fileData)
                {
                    worksheet.Cell(row, 1).Value = data.FileName;
                    worksheet.Cell(row, 4).Value = data.EncryptionTime;
                    worksheet.Cell(row, 3).Value = data.Hash;
                    worksheet.Cell(row, 2).Value = data.FileSize;
                    row++;
                }

                workbook.SaveAs(outputPath);
            }
        }

        public static List<(string FileName, double EncryptionTime, string Hash, long FileSize)> EncryptAndHashDirectorySequentially(string directoryPath, Func<HashAlgorithm> createAlgorithm)
        {
            var results = new List<(string FileName, double EncryptionTime, string Hash, long FileSize)>();
            var files = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                string encryptedFile = $"{file}.enc";
                string hash;
                double duration = EncryptAndHashFile(file, encryptedFile, out hash, createAlgorithm);
                long fileSize = new FileInfo(file).Length;
                results.Add((file, duration, hash, fileSize));
            }

            return results;
        }

        public static string CreateDirectoryIfNotExist(string basePath, string folderName)
        {
            string fullPath = Path.Combine(basePath, folderName);
            if (!Directory.Exists(fullPath))
            {
                Directory.CreateDirectory(fullPath);
                Console.WriteLine($"Created directory at: {fullPath}");
            }
            return fullPath;
        }




        public static List<(string FileName, double EncryptionTime, string Hash)> EncryptAndChCh20DirectorySequentially(string directoryPath, byte[] key, byte[] iv)
        {
            var results = new List<(string FileName, double EncryptionTime, string Hash)>();
            var files = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                byte[] fileData = File.ReadAllBytes(file);
                string hash;
                double duration = StreamCipher.MeasureEncryptionTime(fileData, key, iv, StreamCipher.ChaCha20Encrypt);
                hash = "";
                results.Add((file, duration, hash));
            }

            return results;
        }

        public static async Task<List<(string FileName, double EncryptionTime, string Hash)>> EncryptAndChCh20DirectoryAsync(string directoryPath, byte[] key, byte[] iv)
        {
            var results = new List<(string FileName, double EncryptionTime, string Hash)>();
            var files = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                string encryptedFile = $"{file}.enc";

                var result = await Task.Run(() =>
                {
                    byte[] fileData = File.ReadAllBytes(file);
                    string hash;
                    double duration = StreamCipher.MeasureEncryptionTime(fileData, key, iv, StreamCipher.ChaCha20Encrypt);
                    hash = "";
                    return (FileName: file, EncryptionTime: duration, Hash: hash);
                });

                results.Add(result);
            }

            return results;
        }


        public static void ExcelMergerFiles(List<string> filePaths, Boolean isSequential)
        {



            // Create the master workbook and worksheet
            using (var masterWorkbook = new XLWorkbook())
            {
                var masterWorksheet = masterWorkbook.Worksheets.Add("Merged Data");
                // Assuming "File Name" is the shared column and it's in the first column of each file
                masterWorksheet.Cell(1, 1).Value = "File Name";

                // Dictionary to keep track of file names and corresponding rows
                Dictionary<string, IXLRow> fileNameRowMap = new Dictionary<string, IXLRow>();

                foreach (string path in filePaths)
                {
                    using (var workbook = new XLWorkbook(path))
                    {
                        var worksheet = workbook.Worksheet(1);
                        MapAndMergeData(worksheet, masterWorksheet, fileNameRowMap);
                    }
                }

                masterWorkbook.SaveAs($"{(isSequential ? "Sequinta" : "Paralllel")}merged_file.xlsx");
            }

            Console.WriteLine("Files merged successfully.");
        }
        static void DefineMasterStructure(IXLWorksheet sourceWorksheet, IXLWorksheet masterWorksheet)
        {
            // Copy the headers from the first file
            foreach (var cell in sourceWorksheet.FirstRowUsed().CellsUsed())
            {
                masterWorksheet.Cell(1, cell.Address.ColumnNumber).Value = cell.Value;
            }
        }

        static void MapAndMergeData(IXLWorksheet source, IXLWorksheet target, Dictionary<string, IXLRow> fileNameRowMap)
        {
            // Find the index of the "File Name" column
            int fileNameColumnIndex = source.FirstRowUsed().CellsUsed().First(c => c.Value.ToString() == "File Name").Address.ColumnNumber;

            // Iterate over all rows in the source worksheet
            foreach (var row in source.RowsUsed().Skip(1)) // Skip header row
            {
                string fileName = row.Cell(fileNameColumnIndex).GetString();
                IXLRow targetRow;

                if (fileNameRowMap.ContainsKey(fileName))
                {
                    // If the file name exists, update the existing row
                    targetRow = fileNameRowMap[fileName];
                }
                else
                {
                    // If the file name does not exist, add a new row to the dictionary and worksheet
                    targetRow = target.Row(target.LastRowUsed().RowNumber() + 1);
                    fileNameRowMap.Add(fileName, targetRow);
                }

                // Map all used cells from the source row to the target row
                foreach (var cell in row.CellsUsed())
                {
                    string header = source.Cell(1, cell.Address.ColumnNumber).Value.ToString();
                    // Ensure the target worksheet has this column; if not, create it
                    if (!target.FirstRowUsed().CellsUsed().Any(c => c.Value.ToString() == header))
                    {
                        int newColumn = target.LastColumnUsed().ColumnNumber() + 1;
                        target.Cell(1, newColumn).Value = header;
                    }

                    // Find the column number for this header in the target worksheet
                    int targetColumn = target.FirstRowUsed().CellsUsed().First(c => c.Value.ToString() == header).Address.ColumnNumber;
                    // Set the value of the cell in the target worksheet
                    targetRow.Cell(targetColumn).Value = cell.Value;
                }
            }
        }

        public static void ProcessOfSequentialTechniques(List<int>? filesCounts = null)
        {
            // Your method implementation here
            var stopwatchSequentialSHA256 = Stopwatch.StartNew();
            var sequentialResults = EncryptAndHashDirectorySequentially(directoryPath, () => SHA256.Create());
            stopwatchSequentialSHA256.Stop();
            Console.WriteLine($"Sequential processing  SHA256 time: {stopwatchSequentialSHA256.ElapsedMilliseconds} Milliseconds");
            string sequentialSHA256FolderPath = CreateDirectoryIfNotExist(baseDirectory, "SHA256EncryptionFolderSequential");
            string excelPathSequentialSHA256 = Path.Combine(sequentialSHA256FolderPath, "SHA256EncryptionDataSequential.xlsx");
            WriteToExcel(sequentialResults, excelPathSequentialSHA256, "SHA256", true);

            var stopwatchSequentialSHA512 = Stopwatch.StartNew();
            var sequentialResultsSHA512 = EncryptAndHashDirectorySequentially(directoryPath, () => SHA512.Create());
            stopwatchSequentialSHA512.Stop();
            Console.WriteLine($"Sequential processing  SHA512 time: {stopwatchSequentialSHA512.ElapsedMilliseconds} Milliseconds");
            string sequentialSHA512FolderPath = CreateDirectoryIfNotExist(baseDirectory, "512EncryptionFolderSequential");
            string excelPathSequentialSHA512 = Path.Combine(sequentialSHA512FolderPath, "512EncryptionDataSequential.xlsx");
            WriteToExcel(sequentialResultsSHA512, excelPathSequentialSHA512, "SHA512", true);


            var stopwatchSequentialSHA1 = Stopwatch.StartNew();
            var sequentialResultsSHA1 = EncryptAndHashDirectorySequentially(directoryPath, () => SHA1.Create());
            stopwatchSequentialSHA1.Stop();

            Console.WriteLine($"Sequential processing  SHA1 time: {stopwatchSequentialSHA1.ElapsedMilliseconds} Milliseconds");
            string sequentialSHA1FolderPath = CreateDirectoryIfNotExist(baseDirectory, "SHA1EncryptionFolderSequential");
            string excelPathSequentialSHA1 = Path.Combine(sequentialSHA1FolderPath, "SHA1EncryptionDataSequential.xlsx");
            WriteToExcel(sequentialResultsSHA1, excelPathSequentialSHA1, "SHA1", true);


            var stopwatchSequentialMD5 = Stopwatch.StartNew();
            var sequentialResultsMD5 = EncryptAndHashDirectorySequentially(directoryPath, () => MD5.Create());
            stopwatchSequentialSHA1.Stop();
            Console.WriteLine($"Sequential processing  MD5 time: {stopwatchSequentialMD5.ElapsedMilliseconds} Milliseconds");
            string sequentialMD5FolderPath = CreateDirectoryIfNotExist(baseDirectory, "MD5EncryptionFolderSequential");
            string excelPathSequentialMD5 = Path.Combine(sequentialMD5FolderPath, "MD5EncryptionDataSequential.xlsx");
            WriteToExcel(sequentialResultsMD5, excelPathSequentialMD5, "MD5", true);

            ExcelMergerFiles(new List<string>() { excelPathSequentialSHA256, excelPathSequentialMD5, excelPathSequentialSHA1, excelPathSequentialSHA512 }, true);


        }

        public  static async Task ProcessOfParallelTechniques(List<int>? filesCounts = null)
        {
            var stopwatchSHA256Parallel = Stopwatch.StartNew();
            var parallelResults = await  EncryptAndHashDirectoryAsync(directoryPath, () => SHA256.Create());
            stopwatchSHA256Parallel.Stop();
            Console.WriteLine($"SHA256 Parallel processing time: {stopwatchSHA256Parallel.ElapsedMilliseconds} Milliseconds");
            string parallelSHA256FolderPath = CreateDirectoryIfNotExist(baseDirectory, "SHA256EncryptionFolderParallel");
            string excelPathSHA256Parallel = Path.Combine(parallelSHA256FolderPath, "SHA256EncryptionDataParallel.xlsx");
            WriteToExcel(parallelResults, excelPathSHA256Parallel, "SHA256", false);

            ////var stopwatchMD5Parallel = Stopwatch.StartNew();
            ////var parallelMD5Results = await Utilities.EncryptAndHashDirectoryAsync(directoryPath, () => MD5.Create());
            ////stopwatchMD5Parallel.Stop();
            ////Console.WriteLine($"MD5 Parallel processing time: {stopwatchMD5Parallel.ElapsedMilliseconds} Milliseconds");
            ////string parallelMD5FolderPath = Utilities.CreateDirectoryIfNotExist(baseDirectory, "MD5EncryptionFolderParallel");
            ////string excelPathMD5Parallel = Path.Combine(parallelMD5FolderPath, "MD5EncryptionDataParallel.xlsx");
            ////Utilities.WriteToExcel(parallelMD5Results, excelPathMD5Parallel, "MD5", false);

            ////var stopwatchSHA1Parallel = Stopwatch.StartNew();
            ////var parallelSHA1Results = await Utilities.EncryptAndHashDirectoryAsync(directoryPath, () => SHA1.Create());
            ////stopwatchSHA1Parallel.Stop();
            ////Console.WriteLine($"SHA1 Parallel processing time: {stopwatchSHA1Parallel.ElapsedMilliseconds} Milliseconds");
            ////string parallelSHA1FolderPath = Utilities.CreateDirectoryIfNotExist(baseDirectory, "SHA1EncryptionFolderParallel");
            ////string excelPathSHA1Parallel = Path.Combine(parallelSHA1FolderPath, "SHA1EncryptionDataParallel.xlsx");
            ////Utilities.WriteToExcel(parallelSHA1Results, excelPathSHA1Parallel, "SHA1", false);

            ////var stopwatchSHA512Parallel = Stopwatch.StartNew();
            ////var parallelSHA512Results = await Utilities.EncryptAndHashDirectoryAsync(directoryPath, () => SHA512.Create());
            ////stopwatchSHA512Parallel.Stop();
            ////Console.WriteLine($"SHA512 Parallel processing time: {stopwatchSHA512Parallel.ElapsedMilliseconds} Milliseconds");
            ////string parallelSHA512FolderPath = Utilities.CreateDirectoryIfNotExist(baseDirectory, "SHA512EncryptionFolderParallel");
            ////string excelPathSHA512Parallel = Path.Combine(parallelSHA512FolderPath, "SHA512Encryp.tionDataParallel.xlsx");
            ////Utilities.WriteToExcel(parallelSHA512Results, excelPathSHA512Parallel, "SHA512", false);

            ////Utilities.ExcelMergerFiles(new List<string>() { excelPathSHA256Parallel, excelPathMD5Parallel, excelPathSHA1Parallel, excelPathSHA512Parallel }, false);

        }
    }

}

