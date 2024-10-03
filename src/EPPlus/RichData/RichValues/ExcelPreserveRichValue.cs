using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal class ExcelPreserveRichValue : ExcelRichValue
    {
        public ExcelPreserveRichValue(ExcelWorkbook workbook) : base(workbook, RichDataStructureTypes.Preserve)
        {
        }
    }
}
