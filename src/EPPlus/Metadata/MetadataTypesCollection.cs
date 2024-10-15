using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class MetadataTypesCollection : IndexedCollection<ExcelMetadataType>
    {
        public MetadataTypesCollection(ExcelRichData richData) : base(richData, RichDataEntities.MetadataType)
        {
        }

        public override RichDataEntities EntityType => RichDataEntities.MetadataType;
    }
}
