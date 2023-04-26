using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData.Types;
using System;
using System.Collections.Generic;


namespace OfficeOpenXml.RichData
{
    internal class ExcelRichData
    {
        internal ExcelRichData(ExcelWorkbook wb)
        {
            ValueTypes = new ExcelRichDataValueTypeInfo(wb);
            Structures = new ExcelRichValueStructureCollection(wb);
            Values = new ExcelRichValueCollection(wb, Structures);
        }
        public ExcelRichDataValueTypeInfo ValueTypes { get; }
        public ExcelRichValueStructureCollection Structures { get; }
        public ExcelRichValueCollection Values { get; }
    }
}
