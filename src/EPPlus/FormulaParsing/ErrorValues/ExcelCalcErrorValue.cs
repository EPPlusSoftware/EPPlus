using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    internal class ExcelCalcErrorValue : ExcelRichDataErrorValue
    {
        internal ExcelCalcErrorValue() : base(eErrorType.Calc)
        {
        }
    }
}
