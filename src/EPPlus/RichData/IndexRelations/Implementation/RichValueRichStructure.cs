using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations.Implementation
{
    internal class RichValueRichStructure : IndexPointer
    {
        public RichValueRichStructure(int originalValue) : base(originalValue, RichDataEntities.RichValue, RichDataEntities.RichStructure)
        {
        }
    }
}
