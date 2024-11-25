using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    internal class ExcelBusyErrorValue : ExcelErrorValue
    {
        public ExcelBusyErrorValue()
            : base(eErrorType.Busy)
        {
            
        }
    }
}
