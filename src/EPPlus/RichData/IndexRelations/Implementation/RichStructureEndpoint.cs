using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations.Implementation
{
    internal class RichStructureEndpoint : IndexEndpoint
    {
        public RichStructureEndpoint(ExcelRichData richData) : base(richData.IndexStore, RichDataEntities.RichStructure)
        {
        }
    }
}
