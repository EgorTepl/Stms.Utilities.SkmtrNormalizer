using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Хранит компоненты экземпляра рабочей книги
    /// </summary>
    public class ValidationPropertiesStore
    {
        public SpreadsheetDocument SpreadsheetDocument { get; private set; }
        public WorksheetPart WorksheetPart { get; private set; }
        public string ExcelSheetName { get; private set; }
        public string ExcelFileName { get; private set; }
        public ValidationPropertiesStore(SpreadsheetDocument spreadsheetDocument, Sheet sheet, string fileName)
        {
            SpreadsheetDocument = spreadsheetDocument;
            WorksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id);
            ExcelSheetName = sheet.Name;
            ExcelFileName = fileName;
        }
    }
}
