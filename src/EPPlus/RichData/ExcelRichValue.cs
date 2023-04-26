using System.Collections.Generic;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichValue
    {
        public ExcelRichValue(int structureId)
        {
            StructureId = structureId;
        }

        public int StructureId { get; set; }
        public ExcelRichValueStructure Structure { get; set; }
        public List<object> Values { get; }=new List<object>();
        public RichValueFallbackType Fallback { get; internal set; } = RichValueFallbackType.Decimal;
    }
}