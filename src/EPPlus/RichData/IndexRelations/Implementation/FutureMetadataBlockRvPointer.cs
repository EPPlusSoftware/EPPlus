using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations.Implementation
{
    internal abstract class FutureMetadataBlockRvPointer : IndexPointer
    {
        protected FutureMetadataBlockRvPointer(int originalValue) : base(originalValue, RichDataEntities.FutureMetadataBlock, RichDataEntities.RichValue)
        {
        }
    }
}
