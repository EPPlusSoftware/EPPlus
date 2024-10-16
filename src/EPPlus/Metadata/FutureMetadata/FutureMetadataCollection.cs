using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataCollection : IndexedCollection<ExcelFutureMetadata>
    {
        public FutureMetadataCollection(ExcelRichData richData) : base(richData, RichDataEntities.FutureMetadata)
        {
        }

        public override RichDataEntities EntityType => RichDataEntities.FutureMetadata;
    }
}
