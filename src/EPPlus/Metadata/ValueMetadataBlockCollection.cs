﻿using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class ValueMetadataBlockCollection : IndexedCollection<ExcelValueMetadataBlock>
    {
        public ValueMetadataBlockCollection(ExcelRichData richData) : base(richData, RichDataEntities.ValueMetadataBlock)
        {
        }

        public override RichDataEntities EntityType => RichDataEntities.ValueMetadataBlock;

        public override void Add(ExcelValueMetadataBlock item)
        {
            base.Add(item);
            
        }
    }
}
