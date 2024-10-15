using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations.Implementation
{
    internal class RichValueEndpoint : IndexEndpoint
    {
        public RichValueEndpoint(ExcelRichData richData) : base(richData.IndexStore, RichDataEntities.RichValue)
        {
        }
    }
}
