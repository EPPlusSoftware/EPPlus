using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations.Implementation
{
    internal abstract class FutureMetadataBlockRvEndpoint : IndexEndpoint
    {
        protected FutureMetadataBlockRvEndpoint(ExcelRichData richData) : base(richData.IndexStore, RichDataEntities.FutureMetadataRichDataBlock)
        {
        }
    }
}
