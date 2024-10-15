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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a bk-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelValueMetadataBlock : IndexEndpoint
    {
        public ExcelValueMetadataBlock(ExcelMetadata metadata, RichDataIndexStore store) : base(store, RichDataEntities.ValueMetadataBlock)
        {
            _metadata = metadata;
        }
        public ExcelValueMetadataBlock(XmlReader xr, ExcelMetadata metadata, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataBlock)
        {
            _metadata = metadata;
            while(xr.IsEndElementWithName("bk")==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("rc"))
                {
                    var t = int.Parse(xr.GetAttribute("t"));
                    var v = int.Parse(xr.GetAttribute("v"));
                    Records.Add(new ExcelMetadataRecord(t, v));
                    var metadataType = metadata.MetadataTypes[t - 1];
                    metadata.MetadataTypes.CreateRelation(this, metadataType, IndexType.OneBasedPointer);
                }
                xr.Read();
            }
        }

        private readonly ExcelMetadata _metadata;

        public ExcelValueMetadataBlock(RichDataIndexStore store, RichDataEntities entity) : base(store, entity)
        {
        }

        public List<ExcelMetadataRecord> Records { get;}= new List<ExcelMetadataRecord>();

        public void CreateRelations()
        {
            if(Records != null && Records.Any())
            {
                var pointer = Records.First();
                var relation = _metadata.MetadataTypes.CreateRelation(this, _metadata.MetadataTypes[pointer.TypeIndex - 1], IndexType.OneBasedPointer);
                var item = _metadata.MetadataTypes.GetItem(relation.To.Id);
            }
        }
    }
}