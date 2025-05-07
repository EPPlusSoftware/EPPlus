using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class MdxMetadataCollection : IndexedCollection<MdxMetadata>
    {
        public MdxMetadataCollection(RichDataIndexStore store) : base(store, RichDataEntities.MdxMetadata)
        {
        }
    }
}
