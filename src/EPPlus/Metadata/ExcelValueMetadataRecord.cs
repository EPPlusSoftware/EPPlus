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
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a rc-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelValueMetadataRecord : IndexEndpoint
    {
        public ExcelValueMetadataRecord(MetadataDatabase metadataDb, IndexEndpoint parent, uint typeId, uint valueId, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataRecord)
        {
            TypeId = typeId;
            ValueId = valueId;
            var type = metadataDb.MetadataTypes.Get(typeId);
            var fmt = type.GetFirstOutgoingRelByType<FutureMetadataBase>();
            if(fmt != null)
            {
                var bk = metadataDb.FutureMetadataRichValueBlocks.Get(valueId);
                AddRelationTo(bk);
            }
            AddRelationTo(type);
            _metadataDb = metadataDb;
            _parent = parent;
        }

        public ExcelValueMetadataRecord(XmlReader xr, MetadataDatabase metadataDb, IndexEndpoint parent, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataRecord)
        {
            _metadataDb = metadataDb;
            _parent = parent;
            var t = int.Parse(xr.GetAttribute("t"));
            var v = int.Parse(xr.GetAttribute("v"));
            var type = metadataDb.MetadataTypes[t - 1];
            TypeId = type.Id;
            AddRelationTo(type);
            var fmt = type.GetFirstOutgoingRelByType<FutureMetadataBase>();
            if (fmt != null)
            {
                var bk = fmt.Blocks[v];
                ValueId = bk.Id;
                AddRelationTo(bk);
            }
        }

        private readonly IndexEndpoint _parent;
        private readonly MetadataDatabase _metadataDb;
        //private readonly int _readValueIndex;

        /// <summary>
        /// Corresponds to the t-attribute of the bk element
        /// </summary>
        public uint TypeId { get; private set; }

        /// <summary>
        /// Corresponds to the v-attribute of the bk element
        /// </summary>
        public uint ValueId { get; private set; }

        public int? MetadataTypeIndex
        {
            get
            {
                var ix = _metadataDb.MetadataTypes.GetIndexById(TypeId);
                return ix == null ? null : ix.Value + 1;
            }
        }

        public int? FutureMetadataBlockIndex
        {
            get
            {
                var bk = _metadataDb.FutureMetadataRichValueBlocks.Get(ValueId);
                if (bk == null) return default;
                var fmType = bk.GetFirstIncomingRelByType<FutureMetadataBase>();
                if (fmType == null) return null;
                return fmType.Blocks.GetZeroBasedIndex(ValueId);
            }
        }

        public int? MdxValueMetadataIndex
        {
            get
            {
                var type = _metadataDb.MetadataTypes[(int)TypeId - 1];
                if(type.Name == "XLMDX")
                {
                    return (int)ValueId;
                }
                return null;
            }
        }


        public bool IsValid
        {
            get
            {
                return !Deleted && MetadataTypeIndex.HasValue && MetadataTypeIndex.Value > 0 && (FutureMetadataBlockIndex.HasValue || MdxValueMetadataIndex.HasValue);
            }
        }

        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            base.DeleteMe(relDeletions);
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            DeleteMe(e.RelationDeletions);
        }
    }
}