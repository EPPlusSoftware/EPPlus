using OfficeOpenXml.RichData.Structures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal class ExcelPreserveRichValue : ExcelRichValue
    {
        public ExcelPreserveRichValue(ExcelRichData richData, int structureId, ExcelRichValueStructure structure) : base(richData, structure.StructureType)
        {
            StructureId = structureId;
            Structure = structure;
        }
    }
}
