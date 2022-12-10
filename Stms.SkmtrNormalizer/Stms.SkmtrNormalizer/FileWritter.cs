using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Заполняет выходной Excel файл нормализованными данными
    /// </summary>
    public class FileWritter
    {
        /// <summary>
        ///     Выполняет запись строк в выходной файл
        /// </summary>        
        /// <param name="xlsxRows">Список строк для записи в файл</param>
        /// <param name="rowIndexToWriteFrom">Индекс строки файла, начиная с которой можно выполнять запись</param>
        /// <param name="xlsxFilePath">Путь к файлу для заполнения</param>
        /// <returns>Индекс строки выходного файла, с которой можно выполнять запись следующей порции строк</returns>
        public int WriteToXlsxFile(IList<XlsxOutputTableRowModel> xlsxRows, int rowIndexToWriteFrom, string xlsxFilePath)
        {
            const bool _WRITABLE = true;

            SpreadsheetDocument _spreadsheetDocument = SpreadsheetDocument.Open(xlsxFilePath, _WRITABLE);

            WorkbookPart _workbookPart = _spreadsheetDocument.WorkbookPart;
            WorksheetPart _worksheetPart = _workbookPart.WorksheetParts.First();

            Worksheet _worksheet = _worksheetPart.Worksheet;
            SheetData _sheetData = _worksheet.GetFirstChild<SheetData>();

            rowIndexToWriteFrom = WriteToXlsxSheet(_sheetData, xlsxRows, rowIndexToWriteFrom);            

            _workbookPart.Workbook.Save();
            _spreadsheetDocument.Close();

            return rowIndexToWriteFrom;
        }
        /// <summary>
        ///     Выполняет запись строк на лист выходного файла
        /// </summary>
        /// <param name="sheetData">Контейнер для листов книги</param>
        /// <param name="xlsxRows">Список строк для записи в файл</param>
        /// <param name="rowIndexToWriteFrom">Индекс строки файла, начиная с которой можно выполнять запись</param>
        /// <returns>Индекс строки выходного файла, с которой можно выполнять запись следующей порции строк</returns>
        private int WriteToXlsxSheet(SheetData sheetData, IList<XlsxOutputTableRowModel> xlsxRows, int rowIndexToWriteFrom)
        {
            foreach (var _xlsxRow in xlsxRows)
            {
                var _rowIndexToWriteFrom = 2;
                _rowIndexToWriteFrom += xlsxRows.IndexOf(_xlsxRow) + rowIndexToWriteFrom;

                Row _row = new Row();

                _row.Append(CreateCell(_xlsxRow.SkmtrCode, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.Name, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.Brand, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.StateStandard, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.Size, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.Unit, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.TO2, CellValues.Number));
                _row.Append(CreateCell(_xlsxRow.TO3, CellValues.Number));
                _row.Append(CreateCell(_xlsxRow.TR1, CellValues.Number));
                _row.Append(CreateCell(_xlsxRow.TR2, CellValues.Number));
                _row.Append(CreateCell(_xlsxRow.TR3, CellValues.Number));
                _row.Append(CreateCell(_xlsxRow.SR, CellValues.Number));
                _row.Append(CreateCell(_xlsxRow.Note, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.DaxCode, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.ExcelSheetName, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.ExcelFileName, CellValues.String));
                _row.Append(CreateCell(_xlsxRow.PositionTag, CellValues.String));

                sheetData.Append(_row);
            }
            rowIndexToWriteFrom = xlsxRows.Count + rowIndexToWriteFrom;
            return rowIndexToWriteFrom;
        }
        /// <summary>
        ///     Создает ячейку
        /// </summary>
        /// <param name="cellValue">Значение ячейки</param>
        /// <param name="dataType">Тип данных</param>
        /// <returns>Ячейка</returns>
        private Cell CreateCell(string cellValue, EnumValue<CellValues> dataType)
        {
            Cell _cell = new Cell()
            {
                CellValue = new CellValue(cellValue),
                DataType = dataType
            };
            return _cell;
        }
    }
}
