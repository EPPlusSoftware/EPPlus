using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Types
{
    internal class ExcelRichTypeValueKey
    {
        public ExcelRichTypeValueKey(string name)
        {
            Name = name;
        }
        public string Name { get; set; }
        public RichValueKeyFlags Flags { get; set; }
    }
}
