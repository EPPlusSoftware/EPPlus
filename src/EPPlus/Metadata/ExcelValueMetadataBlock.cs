/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/25/2024         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a rc-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelValueMetadataBlock : IndexEndpoint
    {
        public ExcelValueMetadataBlock(ExcelMetadata metadata, int recordTypeIndex, int valueTypeIndex, ExcelRichData richData)
            : base(richData.IndexStore, RichDataEntities.ValueMetadataRecord)
        {
            var mainRelation = new IndexRelation(this, IndexEndpoint.GetSubRelationsEndpoint(richData.IndexStore), IndexType.SubRelations);
            var record = new ExcelValueMetadataRecord(metadata, this, recordTypeIndex, valueTypeIndex, richData.IndexStore);
            // 1. Add metadata type relation
            var rel1 = new IndexRelation(this, metadata.MetadataTypes[recordTypeIndex], IndexType.OneBasedPointer);
            mainRelation.To.SubRelations.Add(rel1);
            var type = metadata.MetadataTypes.GetItem(rel1.To.Id);
            // 2. Add rich value relation
            var rel2 = richData.Values.CreateRelation(this, valueTypeIndex, IndexType.ZeroBasedPointer);

        }

        public ExcelValueMetadataBlock(XmlReader xr, ExcelMetadata metadata, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataRecord)
        {
            while (xr.IsEndElementWithName("bk") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("rc"))
                {
                    var t = int.Parse(xr.GetAttribute("t"));
                    var v = int.Parse(xr.GetAttribute("v"));
                    Records.Add(new ExcelValueMetadataRecord(metadata, this, t, v, store));
                }
                xr.Read();
            }
        }

        public List<ExcelValueMetadataRecord> Records { get; } = new List<ExcelValueMetadataRecord>();
    }
}