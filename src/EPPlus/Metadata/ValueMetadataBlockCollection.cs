using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class ValueMetadataBlockCollection : IndexedCollection<ExcelValueMetadataBlock>
    {
        public ValueMetadataBlockCollection(RichDataIndexStore store) : base(store, RichDataEntities.ValueMetadataBlock)
        {
        }
    }
}
