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
using ClosedXML.Excel;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GP_EncriptionFiles
{
    public  class StreamCipher
    {
        public static void StreamCipherTest(string folderPath,string  algorithm , TypeGP type= TypeGP.Encription , long filesCounts=100)
        {
           // string folderPath = @"C:\Users\Toshiba\Desktop\Ala'aTheses\GP_EncriptionFiles\bin\Debug\net8.0\download_sample_files";
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Encryption Times");

            worksheet.Cell("A1").Value = "File Name";
            worksheet.Cell("B1").Value = "AES Encryption Time (ms)";
            worksheet.Cell("C1").Value = "AES Decryption Time (ms)";
            worksheet.Cell("D1").Value = "ChaCha20 Encryption Time (ms)";
            worksheet.Cell("E1").Value = "ChaCha20 Decryption Time (ms)";
            switch (algorithm)
            {
                case "chacha":

                    if (type == TypeGP.Encription)
                    {
                        worksheet.Cell("D1").Value = "ChaCha20 Encryption Time (ms)";
                    }
                    else
                    {
                        worksheet.Cell("E1").Value = "ChaCha20 Decryption Time (ms)";
                    }

                    break;
                case "AES":
                    if(type == TypeGP.Encription)
                    {
                        worksheet.Cell("B1").Value = "AES Encryption Time (ms)";
                    }
                    else
                    {
                        worksheet.Cell("C1").Value = "AES Decryption Time (ms)";
                    }
                    break;
                // You can have any number of case statements.
                default:
                    // code to be executed if expression doesn't match any case
                    break;
            }


            var chachaStopwatch = Stopwatch.StartNew();
            int row = 2;
            foreach (string filePathaa in Directory.EnumerateFiles(folderPath))
            {
                byte[] fileData = File.ReadAllBytes(filePathaa);

                // AES Encryption/Decryption
                //  var aesKey = GenerateRandomBytes(32); // 256 bits
                // var aesIV = GenerateRandomBytes(16); // 128 bits
                //var aesEncryptionTime = MeasureEncryptionTime(fileData, aesKey, aesIV, AesEncrypt);
                // var aesDecryptionTime = MeasureEncryptionTime(fileData, aesKey, aesIV, AesDecrypt);
                // ChaCha20 Encryption/Decryption
           
                var chachaKey = GenerateRandomBytes(32); // 256 bits
                var chachaIV = GenerateRandomBytes(12); // 96 bits
                var chachaEncryptionTime = MeasureEncryptionTime(fileData, chachaKey, chachaIV, ChaCha20Encrypt);
                worksheet.Cell(row, 4).Value = chachaEncryptionTime;
                //switch (algorithm)
                //{
                //    case "chacha":

                //        if (type == TypeGP.Encription)
                //        {
                //            var chachaEncryptionTime = MeasureEncryptionTime(fileData, chachaKey, chachaIV, ChaCha20Encrypt);
                //            worksheet.Cell(row, 4).Value = chachaEncryptionTime;
                //        }
                //        else
                //        {
                //            var chachaDecryptionTime = MeasureEncryptionTime(fileData, chachaKey, chachaIV, ChaCha20Decrypt);
                //            worksheet.Cell(row, 5).Value = chachaDecryptionTime;
                //        }

                //        break;
                //    case "AES":
                //        // code to be executed if expression == constant2
                //        break;
                //    // You can have any number of case statements.
                //    default:
                //        // code to be executed if expression doesn't match any case
                //        break;
                //}




                worksheet.Cell(row, 1).Value = Path.GetFileName(filePathaa); 
                //worksheet.Cell(row, 2).Value = aesEncryptionTime;
                //worksheet.Cell(row, 3).Value = aesDecryptionTime;
               
              

                row++;
            }
            chachaStopwatch.Stop();
            Console.WriteLine($"{algorithm} {(type == TypeGP.Encription ? "Encription" : "Decription")} Time: {chachaStopwatch.Elapsed.TotalSeconds} TotalSeconds"); workbook.SaveAs("EncryptionTimesStreem.xlsx"); 
 
        }


         
        // AES encryption
        static byte[] AesEncrypt(byte[] data, byte[] key, byte[] iv)
        {
            using (var aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                var encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
                using (var ms = new MemoryStream())
                using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                {
                    cs.Write(data, 0, data.Length);
                    cs.Close();
                    return ms.ToArray();
                }
            }
        }

        static byte[] AesDecrypt(byte[] data, byte[] key, byte[] iv)
        {
            using (var aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                var decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                using (var ms = new MemoryStream(data))
                using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                using (var msOutput = new MemoryStream())
                {
                    cs.CopyTo(msOutput);
                    return msOutput.ToArray();
                }
            }
        }

        // ChaCha20 encryption
       public  static byte[] ChaCha20Encrypt(byte[] data, byte[] key, byte[] iv)
        {
            IStreamCipher cipher = new ChaCha7539Engine();
            cipher.Init(true, new ParametersWithIV(new KeyParameter(key), iv));
            byte[] output = new byte[data.Length];
            cipher.ProcessBytes(data, 0, data.Length, output, 0);
            return output;
        }

        // ChaCha20 decryption
        static byte[] ChaCha20Decrypt(byte[] data, byte[] key, byte[] iv)
        {
            IStreamCipher cipher = new ChaCha7539Engine();
            cipher.Init(false, new ParametersWithIV(new KeyParameter(key), iv));
            byte[] output = new byte[data.Length];
            cipher.ProcessBytes(data, 0, data.Length, output, 0);
            return output;
        }

        public static long MeasureEncryptionTime(byte[] data, byte[] key, byte[] iv, Func<byte[], byte[], byte[], byte[]> encryptOrDecryptMethod)
        {
            var stopwatch = Stopwatch.StartNew();
            encryptOrDecryptMethod(data, key, iv);
            stopwatch.Stop();
            return stopwatch.ElapsedMilliseconds;
        }

        public static byte[] GenerateRandomBytes(int length)
        {
            var rng = new RNGCryptoServiceProvider();
            var randomBytes = new byte[length];
            rng.GetBytes(randomBytes);
            return randomBytes;
        }

    }



}
