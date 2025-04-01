using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities.TextUtils
{
    internal class DelimiterInfo
    {
        public DelimiterInfo(int length, int position)
        {
            Length = length;
            Position = position;
        }

        public int Length { get; }
        public int Position { get; set; }
    }
}
