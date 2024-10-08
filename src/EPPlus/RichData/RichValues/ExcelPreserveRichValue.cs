using OfficeOpenXml.RichData.Structures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal class ExcelPreserveRichValue : ExcelRichValue
    {
        public ExcelPreserveRichValue(ExcelRichData richData) : base(richData, RichDataStructureTypes.Preserve)
        {
        }
    }
}
