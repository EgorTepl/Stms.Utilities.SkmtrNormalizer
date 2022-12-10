using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Collections;
using System.IO;
using System.Globalization;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Реализует чтение входных файлов Excel
    /// </summary>
    public class XlsxReader
    {
        private const int INDEX_OF_START_READING_ROWS = 5;

        public IDictionary<string, int> columnLettersLookup { get; private set; }
        public IDictionary<string, int> columnNamesLookup { get; private set; }
        private Regex skmtrRegex;
        private Regex lettersRegex;
        private Regex materialsRegex;
        private Regex repairPartsRegex;

        /// <summary>
        ///     Вызывает метод инициализации переменных
        /// </summary>
        public XlsxReader()
        {
            this.Initialize();
        }
        /// <summary>
        ///     Инициализирует переменные
        /// </summary>
        private void Initialize()
        {
            skmtrRegex = new Regex("[0-9]+");
            lettersRegex = new Regex("[A-Za-z]+");
            materialsRegex = new Regex(@"\b(МАТ|матер|материалы)\b");
            repairPartsRegex = new Regex(@"\b(ЗЧ|запч|зап.части)\b");

            columnLettersLookup = this.CreateColumnLettersLookup();
            columnNamesLookup = this.CreateColumnNamesLookup();
        }
        /// <summary>
        ///     Читает входной Excel-файл 
        /// </summary>
        /// <param name="xlsxFilePath">Путь к файлу Excel</param>
        /// /// <param name="globalSkmtrLookup">Глобальная таблица кодов СК-МТР</param>
        /// <returns>Список заполненных объектов класса строковой модели</returns>
        public IList<XlsxOutputTableRowModel> ParseFile(string xlsxFilePath, HashSet<string> globalSkmtrLookup)
        {
            const bool _READ_ONLY = false;
         
            var _xlsxFileRows = new List<XlsxOutputTableRowModel>();
            var _xlsxFileName = Path.GetFileName(xlsxFilePath);

            using (SpreadsheetDocument _spreadsheetDocument = SpreadsheetDocument.Open(xlsxFilePath, _READ_ONLY))
            {
                WorkbookPart _workbookPart = _spreadsheetDocument.WorkbookPart;                

                foreach (Sheet _sheet in _workbookPart.Workbook.Sheets)
                {
                    if(_sheet.State != null && _sheet.State.HasValue && (_sheet.State.Value == SheetStateValues.Hidden || _sheet.State.Value == SheetStateValues.VeryHidden))
                    {
                        continue;
                    }
 
                    ValidationPropertiesStore _validationPropertiesStore = GetPropertiesForValidation(_spreadsheetDocument, _sheet, _xlsxFileName);

                    SkmtrValidator.EnsureTableHeaderFormatIsValid(_validationPropertiesStore);

                    var _xlsxSheetNamePrefix = GetXlsxSheetNamePrefix(_sheet.Name);

                    foreach (var _row in GetRowsWithSkmtrCode(_validationPropertiesStore))
                    {
                        var _xlsxOutputTableModel = CreateXlsxOutputTableModel(_row, _validationPropertiesStore, _xlsxSheetNamePrefix);
                        var _skmtrCode = string.Format("{0}_{1}", _xlsxSheetNamePrefix, _xlsxOutputTableModel.SkmtrCode);

                        if (globalSkmtrLookup.Contains(_skmtrCode))
                        {
                            continue;
                        }
                        globalSkmtrLookup.Add(_skmtrCode);
                        _xlsxFileRows.Add(_xlsxOutputTableModel);
                    }
                }
            }
            return _xlsxFileRows;
        }
        /// <summary>
        ///     Заполняет объект класса модели строки таблицы
        /// </summary>
        /// <param name="row">Строка листа</param>
        /// /// <param name="sheetName">Имя листа</param>
        /// /// <param name="fileName">Имя файла</param>
        /// <param name="spreadsheetDocument">Компонент табличного документа</param>
        /// <returns>Объект класса табличной модели</returns>
        private XlsxOutputTableRowModel CreateXlsxOutputTableModel(Row row, ValidationPropertiesStore validationPropertiesStore, string xlsxSheetNamePrefix)
        {
            var _xlsxRows = new List<string>();

            foreach (var _cell in GetCellForRow(row, this.columnLettersLookup))
            {
                 _xlsxRows.Add(GetCellValue(validationPropertiesStore.SpreadsheetDocument, _cell));
            }

            var xlsxOutputTableModel = new XlsxOutputTableRowModel()
            {
                SkmtrCode = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.SKMTR_CODE]],
                Name = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.UNIT_NAME]],
                Brand = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.BRAND]],
                StateStandard = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.STATE_STANDARD]],
                Size = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.SIZE]],
                Unit = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.UNIT]],
                TO2 = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.TO_2]],
                TO3 = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.TO_3]],
                TR1 = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.TR_1]],
                TR2 = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.TR_2]],
                TR3 = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.TR_3]],
                SR = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.SR]],
                Note = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.NOTE]],
                DaxCode = _xlsxRows[this.columnNamesLookup[XlsxTableHeader.DAX_CODE]],
                ExcelSheetName = validationPropertiesStore.ExcelSheetName,
                ExcelFileName = validationPropertiesStore.ExcelFileName,
                PositionTag = xlsxSheetNamePrefix
            };
            return xlsxOutputTableModel;
        }
        /// <summary>
        ///     Возвращает префикс имени листа
        /// </summary>
        /// <param name="xlsxSheetName">Имя листа</param>
        /// <returns>Префикс имени листа</returns>
        private string GetXlsxSheetNamePrefix(string xlsxSheetName)
        {
            const string _MATERIALS = "Материалы";
            const string _REPAIR_PARTS = "Запчасти";
            const string _NOT_DEFINED = "Неопределено";

            Match _match = this.materialsRegex.Match(xlsxSheetName);

            if (_match.Success)
            {
                return _MATERIALS;
            }

            _match = this.repairPartsRegex.Match(xlsxSheetName);

            if (_match.Success)
            {
                return _REPAIR_PARTS;
            }

            return _NOT_DEFINED;            
        }
        /// <summary>
        ///     Возвращает компоненты экзмепляра рабочей книги
        /// </summary>
        /// <param name="spreadsheetDocument">Компонент табличного документа</param>
        /// <param name="sheet">Лист</param>
        /// <param name="fileName">Имя файла</param>
        /// <returns>Возвращает объект класса свойств для валидации заголовка</returns>
        private ValidationPropertiesStore GetPropertiesForValidation(SpreadsheetDocument spreadsheetDocument, Sheet sheet, string fileName)
        {
            ValidationPropertiesStore _validationPropertiesStore = new ValidationPropertiesStore(spreadsheetDocument, sheet, fileName);

            return _validationPropertiesStore;
        }
        /// <summary>
        ///     Возвращает строку с кодом СК-МТР
        /// </summary>
        /// <param name="spreadsheetDocument">Компонент табличного документа</param>
        /// <param name="worksheetPart">Компонент листа</param>
        /// <returns>Строка</returns>
        private IEnumerable<Row> GetRowsWithSkmtrCode(ValidationPropertiesStore _validationPropertiesStore)
        {
            foreach(var _row in _validationPropertiesStore.WorksheetPart.Worksheet.Descendants<Row>().Skip(XlsxReader.INDEX_OF_START_READING_ROWS))
            {
                if (IsRowShouldIgnored(_row, _validationPropertiesStore.SpreadsheetDocument))
                {
                    continue;
                }

                yield return _row;
            }
        }
        /// <summary>
        ///     Проверяет ячейку на содержание кода СК-МТР
        /// </summary>
        /// <param name="row">Строка листа</param>
        /// <param name="spreadsheetDocument">Компонент табличного документа</param>
        /// <returns></returns>
        private bool IsRowShouldIgnored(Row row, SpreadsheetDocument spreadsheetDocument)
        {
            foreach (var _cell in row.Descendants<Cell>())
            {
                var _hasNoSkmtrCode = string.IsNullOrEmpty(GetSkmtrCode(GetCellValue(spreadsheetDocument, _cell)));
                if (_cell == row.FirstChild && _hasNoSkmtrCode)
                {
                    return true;
                }
            }

            return false;
        }
        /// <summary>
        ///     Проверяет ссылку ячейки на вхождение ее в диапазон заголовка таблицы
        /// </summary>
        /// <param name="row">Строка</param>
        /// <param name="columnLettersLookup">Словарь букв диапазона заголовка</param>
        /// <returns>Ячейка</returns>
        private IEnumerable<Cell> GetCellForRow(Row row, IDictionary<string, int> columnLettersLookup)
        {
            int _cellLetterIndex = 0;

            foreach(var _cell in row.Descendants<Cell>())
            {
                var _cellLetter = GetColumnAddress(_cell.CellReference);

                if (!columnLettersLookup.ContainsKey(_cellLetter))
                {
                    break;
                }

                int _currentCellLetterIndex = columnLettersLookup[_cellLetter];

                for (; _cellLetterIndex < _currentCellLetterIndex; _cellLetterIndex++)
                {
                    var _emptyCell = new Cell { DataType = null, CellValue = new CellValue(string.Empty) };
                    yield return _emptyCell;
                }
                yield return _cell;
                _cellLetterIndex++;

                if (_cell == row.LastChild)
                {
                    for (; _cellLetterIndex < columnLettersLookup.Count(); _cellLetterIndex++)
                    {
                        var _emptyCell = new Cell() { DataType = null, CellValue = new CellValue(string.Empty) };
                        yield return _emptyCell;
                    }
                }
            }
        }
        /// <summary>
        ///     Форматирует ссылку на ячейку
        /// </summary>
        /// <param name="cellReference">Адрес ячейки</param>
        /// <returns>Букву ссылки на ячейку</returns>
        private string GetColumnAddress(string cellReference)
        {
            Match _match = this.lettersRegex.Match(cellReference);
            return _match.Value;
        }
        /// <summary>
        ///     Возвращает значение кода СК-МТР
        /// </summary>
        /// <param name="cellValue">Значение ячейки</param>
        /// <returns>Значение кода СК-МТР</returns>
        private string GetSkmtrCode(string cellValue)
        {
            Match _match = this.skmtrRegex.Match(cellValue);
            return _match.Value;
        }
        /// <summary>
        ///     Создает словарь имен столбцов заголовка обрабатываемого листа
        /// </summary>
        /// <returns>Словарь имен столбцов заголовка</returns>
        public IDictionary<string, int> CreateColumnNamesLookup()
        {
            var _columnNamesLookup = new Dictionary<string, int>();

            _columnNamesLookup.Add(XlsxTableHeader.SKMTR_CODE, 0);
            _columnNamesLookup.Add(XlsxTableHeader.UNIT_NAME, 1);
            _columnNamesLookup.Add(XlsxTableHeader.BRAND, 2);
            _columnNamesLookup.Add(XlsxTableHeader.STATE_STANDARD, 3);
            _columnNamesLookup.Add(XlsxTableHeader.SIZE, 4);
            _columnNamesLookup.Add(XlsxTableHeader.UNIT, 5);
            _columnNamesLookup.Add(XlsxTableHeader.TO_2, 6);
            _columnNamesLookup.Add(XlsxTableHeader.TO_3, 7);
            _columnNamesLookup.Add(XlsxTableHeader.TR_1, 8);
            _columnNamesLookup.Add(XlsxTableHeader.TR_2, 9);
            _columnNamesLookup.Add(XlsxTableHeader.TR_3, 10);
            _columnNamesLookup.Add(XlsxTableHeader.SR, 11);
            _columnNamesLookup.Add(XlsxTableHeader.NOTE, 12);
            _columnNamesLookup.Add(XlsxTableHeader.DAX_CODE, 13);
            _columnNamesLookup.Add(XlsxTableHeader.EXCEL_SHEET_NAME, 14);
            _columnNamesLookup.Add(XlsxTableHeader.EXCEL_FILE_NAME, 15);
            _columnNamesLookup.Add(XlsxTableHeader.POSITION_TAG, 16);

            return _columnNamesLookup;
        }
        /// <summary>
        ///     Создает словарь ссылок на ячейки, входящие в диапазон шаблонного заголовка
        /// </summary>
        /// <returns>Словарь ссылок на ячейки</returns>
        private IDictionary<string, int> CreateColumnLettersLookup()
        {
            var _columnLettersLookup = new Dictionary<string, int>();

            _columnLettersLookup.Add("A", 0);
            _columnLettersLookup.Add("B", 1);
            _columnLettersLookup.Add("C", 2);
            _columnLettersLookup.Add("D", 3);
            _columnLettersLookup.Add("E", 4);
            _columnLettersLookup.Add("F", 5);
            _columnLettersLookup.Add("G", 6);
            _columnLettersLookup.Add("H", 7);
            _columnLettersLookup.Add("I", 8);
            _columnLettersLookup.Add("J", 9);
            _columnLettersLookup.Add("K", 10);
            _columnLettersLookup.Add("L", 11);
            _columnLettersLookup.Add("M", 12);
            _columnLettersLookup.Add("N", 13);

            return _columnLettersLookup;
        }
        /// <summary>
        ///     Возвращает содержимое ячейки
        /// </summary>
        /// <param name="spreadsheetDocument">Компонент табличного документа</param>
        /// <param name="cell">Ячейка</param>
        /// <returns>Содержимое ячейки</returns>
        private string GetCellValue(SpreadsheetDocument spreadsheetDocument, Cell cell)
        {
            var _cellValue = GetValidCellValue(cell);

            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        var _sstPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        _cellValue = _sstPart.SharedStringTable.ChildElements[int.Parse(_cellValue)].InnerText.Trim();
                        break;
                }
            }
            return _cellValue;
        }
        /// <summary>
        ///     Возвращает значение ячейки в зависимости от типа данных ячейки
        /// </summary>
        /// <param name="cell">Ячейка</param>
        /// <returns>Если в ячейке содержится формула, то вернет результат формулы, в противном случае индекс для таблицы строк SharedStringTable</returns>
        private string GetValidCellValue(Cell cell)
        {
            if (cell == null) return null;

            var _cellValue = cell.CellFormula == null ? cell.InnerText : cell.CellValue.InnerText;

            return _cellValue;
        }
    }
}
