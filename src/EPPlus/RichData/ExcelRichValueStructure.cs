using System.Collections.Generic;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichValueStructure
    {
        public string Type { get; set; }
        public List<ExcelRichValueStructureKey> Keys { get;  }=new List<ExcelRichValueStructureKey>();
    }
}
