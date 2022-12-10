using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Reflection;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Создает выходной Excel-файл с расширением .xlsx
    /// </summary>
    public class XlsxCreator
    {
        /// <summary>
        ///     Создает выходной Excel-файл, добавляет в него лист и заголовок будущей таблицы
        /// </summary>
        /// <param name="excelFolderPath">Путь к папке, в которой будет создан файл</param>
        /// <returns>Возращает ссылку на созданный файл</returns>
        public string CreateNewXlsxFile(string excelFolderPath, XlsxReader xlsxReader)
        {
            string _dateTimeNow = DateTime.Now.ToString("yyyyMMddHHmmss");
            string _excelFileName = "Output";
            string _excelFilePath = Path.Combine(excelFolderPath, $"{_excelFileName}_{_dateTimeNow}.xlsx");

            SpreadsheetDocument _spreadsheetDocument = SpreadsheetDocument.Create(_excelFilePath, SpreadsheetDocumentType.Workbook);            
            WorkbookPart _workbookPart = _spreadsheetDocument.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();

            WorksheetPart _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            _worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets _sheets = _spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Sheet _sheet = new Sheet()
            {
                Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(_worksheetPart),
                SheetId = 1,
                Name = "Лист1"
            };
            _sheets.Append(_sheet);

            Worksheet _worksheet = _worksheetPart.Worksheet;
            SheetData _sheetData = _worksheet.GetFirstChild<SheetData>();

            _sheetData.Append(CreateHeaderOnASheet(xlsxReader));

            _workbookPart.Workbook.Save();
            _spreadsheetDocument.Close();
            return _excelFilePath;
        }
        /// <summary>
        ///     Создает заголовок таблицы
        /// </summary>
        /// <returns>Заголовок таблицы</returns>
        private Row CreateHeaderOnASheet(XlsxReader xlsxReader)
        {
            var _columnNamesLookup = xlsxReader.columnNamesLookup;

            Row _row = new Row();

            foreach (var _columnName in _columnNamesLookup.Keys)
            {
                Cell _cell = new Cell()
                {
                    CellValue = new CellValue(_columnName),
                    DataType = CellValues.String
                };
                _row.Append(_cell);
            }
            return _row;
        }
    }
}
