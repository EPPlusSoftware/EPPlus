using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class ExcelColumnComparer : IComparer<object>
    {
        public int Compare(object x, object y)
        {
            throw new NotImplementedException();
        }
        public int Compare(int column, ExcelColumn col)
        {
            throw new NotImplementedException();
        }
    }
}
