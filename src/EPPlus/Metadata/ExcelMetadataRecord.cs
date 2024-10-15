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

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a rc-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelMetadataRecord : IndexEndpoint
    {
        public ExcelMetadataRecord(ExcelMetadata metadata, IndexEndpoint parent, int recordTypeIndex, int valueTypeIndex, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataRecord)
        {
            TypeIndex= recordTypeIndex;
            ValueIndex = valueTypeIndex;
            // 1. Add metadata type relation
            var rel1 = new IndexRelation(this, metadata.MetadataTypes[TypeIndex - 1], IndexType.OneBasedPointer);
            store.AddRelation(rel1);
            parent.SubRelations.Add(rel1);
            var type = metadata.MetadataTypes.GetItem(rel1.To.Id);
        }

        /// <summary>
        /// Corresponds to the t-attribute of the bk element
        /// </summary>
        public int TypeIndex { get; private set; }

        /// <summary>
        /// Corresponds to the v-attribute of the bk element
        /// </summary>
        public int ValueIndex { get; private set; }
    }
}