using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Reflection;
using System.IO;


namespace Stms.SkmtrNormalizer
{
    class Program
    {
        static void Main(string[] args)
        {
            var _excelFolderPath = Path.GetDirectoryName(AppContext.BaseDirectory);

            SkmtrNormalizer _skmtrNormalizer = new SkmtrNormalizer();
            try
            {
                Console.WriteLine(SR.GetDataFromExcelException);
                Console.WriteLine("");
                
                _skmtrNormalizer.NormalizeSkmtrData(_excelFolderPath);

                Console.WriteLine("");
                Console.WriteLine(string.Format(SR.OutputXlsxFileReadyException, _skmtrNormalizer.XlsxOutputFilePath));
                Console.ReadKey();
            }
            catch(Exception exception)
            {
                File.Delete(_skmtrNormalizer.XlsxOutputFilePath);
                Console.WriteLine(exception.Message);
                Console.ReadKey();
            }
        }
    }
}
