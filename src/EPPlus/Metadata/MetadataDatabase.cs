/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  26/11/2024         EPPlus Software AB       EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Metadata.FutureMetadata;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class MetadataDatabase
    {
        public MetadataDatabase(RichDataIndexStore store)
        {
            IndexStore = store;
            ValueMetadata = new ValueMetadataBlockCollection(store);
            ValueMetadataRecords = new ValueMetadataRecordCollection(store);
            CellMetadata = new CellMetadataBlockCollection(store);
            CellMetadataRecords = new CellMetadataRecordCollection(store);
            MetadataTypes = new MetadataTypesCollection(store);
            FutureMetadata = new FutureMetadataCollection(store);
            FutureMetadataRichValueBlocks = new FutureMetadataRichValueBlockCollection(store);
            FutureMetadataDynamicArrayBlocks = new FutureMetadataDynamicArrayBlockCollection(store);
        }

        internal RichDataIndexStore IndexStore { get; private set; }

        internal MetadataTypesCollection MetadataTypes { get; }
        internal FutureMetadataCollection FutureMetadata { get; set; }
        internal ValueMetadataBlockCollection ValueMetadata { get; }

        internal ValueMetadataRecordCollection ValueMetadataRecords { get; }

        internal CellMetadataBlockCollection CellMetadata { get; }

        internal CellMetadataRecordCollection CellMetadataRecords { get; set; }

        internal FutureMetadataRichValueBlockCollection FutureMetadataRichValueBlocks { get; }

        internal FutureMetadataDynamicArrayBlockCollection FutureMetadataDynamicArrayBlocks { get; }
    }
}
