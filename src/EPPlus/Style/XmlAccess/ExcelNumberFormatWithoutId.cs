using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style.Interfaces;

namespace OfficeOpenXml.Style.XmlAccess
{
    internal class ExcelNumberFormatWithoutId : IExcelNumberFormat
    {
        internal ExcelNumberFormatWithoutId(string format) 
        {
            Format = format;
            NumFmtId = -1;
            BuildIn = false;
        }
        public string Format { get; private set; }

        public int NumFmtId { get; private set; }

        public bool BuildIn { get; private set; }
    }
}
