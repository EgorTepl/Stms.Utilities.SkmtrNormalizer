using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Stms.SkmtrNormalizer
{
    /// <summary>
    ///     Реализует установку и считывание строк входного листа файла Excel
    /// </summary>
    public class XlsxOutputTableRowModel
    {
        public string SkmtrCode { get; set; }
        public string Name { get; set; }
        public string Brand { get; set; }
        public string StateStandard { get; set; }
        public string Size { get; set; }
        public string Unit { get; set; }
        public string TO2 {get;set;}
        public string TO3 { get; set; }
        public string TR1 { get; set; }
        public string TR2 { get; set; }
        public string TR3 { get; set; }
        public string SR { get; set; }
        public string Note { get; set; }
        public string DaxCode { get; set; }
        public string ExcelSheetName { get; set; }
        public string ExcelFileName { get; set; }
        public string PositionTag { get; set; }
    }
}
