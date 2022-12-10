using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Содержит наименования столбцов выходной таблицы
    /// </summary>
    public class XlsxTableHeader
    {
        public const string SKMTR_CODE = "Код СК-МТР";
        public const string UNIT_NAME = "Наименование";
        public const string BRAND = "Марка, обозначение чертежа";
        public const string STATE_STANDARD = "ГОСТ, ТУ";
        public const string SIZE = "Сорт, размер";
        public const string UNIT = "Ед. изм.";
        public const string CONSUMPTION_RATE = "Норма расхода";
        public const string TO_2 = "ТО-2";
        public const string TO_3 = "ТО-3";
        public const string TR_1 = "ТР-1";
        public const string TR_2 = "ТР-2";
        public const string TR_3 = "ТР-3";
        public const string SR = "СР";
        public const string NOTE = "Примечание";
        public const string DAX_CODE = "КОД DAX";
        public const string EXCEL_SHEET_NAME = "Имя листа Excel";
        public const string EXCEL_FILE_NAME = "Имя файла Excel";
        public const string POSITION_TAG = "Признак позиции";
    }
}