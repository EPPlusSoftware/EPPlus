using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal class ExcelPreserveRichValue : ExcelRichValue
    {
        public ExcelPreserveRichValue(RichDataIndexStore store, ExcelRichData richData, uint structureId, ExcelRichValueStructure structure)
            : base(store, richData, structure.StructureType)
        {
            StructureId = structureId;
            Structure = structure;
        }
    }
}
