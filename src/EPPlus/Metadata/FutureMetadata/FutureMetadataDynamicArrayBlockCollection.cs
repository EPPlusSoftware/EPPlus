using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataDynamicArrayBlockCollection : IndexedCollection<FutureMetadataBlock>
    {
        public FutureMetadataDynamicArrayBlockCollection(RichDataIndexStore store) : base(store, RichDataEntities.FutureMetadataDynamicArrayBlock)
        {
        }

        public override RichDataEntities EntityType => RichDataEntities.FutureMetadataDynamicArrayBlock;
    }
}
