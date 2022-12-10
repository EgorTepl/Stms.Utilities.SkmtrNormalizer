using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Обрабатывает входные файлы Excel
    /// </summary>
    public class SkmtrNormalizer
    {
        public string XlsxOutputFilePath { get; private set; }
        /// <summary>
        ///     Записывает в выходной Excel-файл нормализованные данные из входных Excel-файлов
        /// </summary>
        /// <param name="excelFolderPath">Путь к папке с файлами Excel</param>
        public void NormalizeSkmtrData(string excelFolderPath)
        {            
            int _rowIndexToWriteFrom = 0;
            var xlsxOutputFolderPath = Path.Combine(excelFolderPath, "Output");
            HashSet<string> globalSkmtrLookup = new HashSet<string>();

            XlsxReader _xlsxReader = new XlsxReader();
            XlsxCreator _xlsxCreator = new XlsxCreator();
            FileWritter _fileWritter = new FileWritter();

            Directory.CreateDirectory(xlsxOutputFolderPath);

            this.XlsxOutputFilePath = _xlsxCreator.CreateNewXlsxFile(xlsxOutputFolderPath, _xlsxReader);

            foreach (var _file in Directory.EnumerateFiles(excelFolderPath, "*.xlsx"))
            {               
                IList<XlsxOutputTableRowModel> _xlsxRows = _xlsxReader.ParseFile(_file, globalSkmtrLookup);

                _rowIndexToWriteFrom = _fileWritter.WriteToXlsxFile(_xlsxRows, _rowIndexToWriteFrom, this.XlsxOutputFilePath);
            }
        }
    }
}
