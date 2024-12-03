using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations.Implementation
{
    internal class RichStructureEndpoint : IndexEndpoint
    {
        public RichStructureEndpoint(RichDataIndexStore store) : base(store, RichDataEntities.RichStructure)
        {
        }
    }
}
