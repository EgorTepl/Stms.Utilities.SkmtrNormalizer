using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Реализует валидацию заголовка таблицы входного листа Excel
    /// </summary>
    public static class SkmtrValidator
    {
        /// <summary>
        ///     Запускает проверку имен столбцов
        /// </summary>
        /// <param name="validationPropertiesStore">Компоненты экземпляра рабочей книги</param>
        public static void EnsureTableHeaderFormatIsValid(ValidationPropertiesStore validationPropertiesStore)
        {
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.SKMTR_CODE, XlsxCellReference.A_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.UNIT_NAME, XlsxCellReference.B_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.BRAND, XlsxCellReference.C_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.STATE_STANDARD, XlsxCellReference.D_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.SIZE, XlsxCellReference.E_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.UNIT, XlsxCellReference.F_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.CONSUMPTION_RATE, XlsxCellReference.G_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.TO_2, XlsxCellReference.G_5);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.TO_3, XlsxCellReference.H_5);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.TR_1, XlsxCellReference.I_5);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.TR_2, XlsxCellReference.J_5);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.TR_3, XlsxCellReference.K_5);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.SR, XlsxCellReference.L_5);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.NOTE, XlsxCellReference.M_4);
            EnsureColumnNameIsValid(validationPropertiesStore, XlsxTableHeader.DAX_CODE, XlsxCellReference.N_4);
        }
        /// <summary>
        ///     Проверяет содержимое ячейки по ссылке со значением из шаблона
        /// </summary>
        /// <param name="validationPropertiesStore">Компоненты экземпляра рабочей книги</param>
        /// <param name="columnNameByPattern">Имя столбца заголовка по шаблону</param>
        /// <param name="cellAddress">Ссылка на ячейку</param>
        private static void EnsureColumnNameIsValid(ValidationPropertiesStore validationPropertiesStore, string columnNameByPattern ,string cellAddress)
        {
            string _sheetName = validationPropertiesStore.ExcelSheetName;
            string _fileName = validationPropertiesStore.ExcelFileName;

            var _cellValue = GetCellValueFromAdress(validationPropertiesStore, cellAddress);            

            if (!columnNameByPattern.StartsWith(_cellValue) || string.IsNullOrEmpty(_cellValue))
            {
                throw new Exception(string.Format(SR.InvalidColumnNameException, cellAddress, columnNameByPattern, _cellValue, _fileName, _sheetName));
            }
        }
        /// <summary>
        ///     Возвращает значение ячейки
        /// </summary>
        /// <param name="validationPropertiesStore">Компоненты экземпляра рабочей книги</param>
        /// <param name="cellAdressName">Ссылка на ячейку</param>
        /// <returns>Значение ячейки</returns>
        private static string GetCellValueFromAdress(ValidationPropertiesStore validationPropertiesStore, string cellAdressName)
        {
            string _cellValue = string.Empty;

            SpreadsheetDocument _spreadsheetDocument = validationPropertiesStore.SpreadsheetDocument;
            WorksheetPart _worksheetPart = validationPropertiesStore.WorksheetPart;

            Cell _cell = _worksheetPart.Worksheet.Descendants<Cell>().
                Where(c => c.CellReference == cellAdressName).FirstOrDefault();

            if(_cell != null)
            {
                _cellValue = _cell.InnerText;
                
                if (_cell.DataType != null)
                { 
                    var _stringTable = _spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (_stringTable != null)
                    {
                        _cellValue = _stringTable.SharedStringTable.ElementAt(int.Parse(_cellValue)).InnerText;
                    }
                }
            }
            return _cellValue = Regex.Replace(_cellValue, @"\s+", " ").Trim();
        }
    }
}
