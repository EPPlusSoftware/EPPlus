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
using OfficeOpenXml.Metadata.FutureMetadata;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using System;
using System.Linq;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a rc-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelValueMetadataRecord : IndexEndpoint
    {
        public ExcelValueMetadataRecord(ExcelMetadata metadata, IndexEndpoint parent, uint typeId, uint valueId, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataRecord)
        {
            TypeId = typeId;
            ValueId = valueId;
            _metadata = metadata;
            _readValueIndex = Convert.ToInt32(valueId);
            _parent = parent;
        }

        private readonly IndexEndpoint _parent;
        private readonly ExcelMetadata _metadata;
        private readonly int _readValueIndex;

        public void InitRelations(ExcelRichData richData)
        {
            base.InitRelations();
            var parentRel = _parent.GetOutgoingRelations(x => x.IndexType == IndexType.SubRelations && x.AsRelationWithSubRelations().SubRelationEntity == RichDataEntities.RichValue).FirstOrDefault();
            if(parentRel != null)
            {
                var rel = richData.Values.CreateRelation(this, _readValueIndex, IndexType.ZeroBasedPointer);
                ValueId = rel.To.Id;
            }
        }

        /// <summary>
        /// Corresponds to the t-attribute of the bk element
        /// </summary>
        public uint TypeId { get; private set; }

        /// <summary>
        /// Corresponds to the v-attribute of the bk element
        /// </summary>
        public uint ValueId { get; private set; }

        public int MetadataTypeIndex
        {
            get
            {
                var ix = _metadata.MetadataTypes.GetIndexById(TypeId);
                return ix.Value + 1;
            }
        }

        public int FutureMetadataBlockIndex
        {
            get
            {
                var bk = _metadata.FutureMetadataBlocks.Get(ValueId);
                var parentRel = _parent.GetOutgoingRelations(x => x.IndexType == IndexType.SubRelations && x.AsRelationWithSubRelations().SubRelationEntity == RichDataEntities.RichValue).FirstOrDefault();
                if(parentRel != null)
                {
                    var parentRelSub = parentRel.AsRelationWithSubRelations();
                    for (var ix = 0; ix < parentRelSub.SubRelations.Count; ix++)
                    {
                        var sr = parentRelSub.SubRelations[ix];
                        if(sr.To.Id == ValueId)
                        {
                            return ix;
                        }
                    }
                }
                return -1;
            }
        }

        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            base.DeleteMe(relDeletions);
            var parent = _parent as ExcelValueMetadataBlock;
            if(parent != null)
            {
                parent.OnRecordDeleted(this, relDeletions);
            }
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            base.OnConnectedEntityDeleted(e);
            _parent.OnConnectedEntityDeleted(e);
        }
    }
}