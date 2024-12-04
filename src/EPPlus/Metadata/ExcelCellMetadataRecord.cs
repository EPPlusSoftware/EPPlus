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
using OfficeOpenXml.Utils;
using System;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a rc-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelCellMetadataRecord : IndexEndpoint
    {
        //public ExcelCellMetadataRecord(MetadataDatabase metadataDb, IndexEndpoint parent, uint typeId, uint valueId, RichDataIndexStore store)
        //    : base(store, RichDataEntities.CellMetadataRecord)
        //{
        //    TypeId = typeId;
        //    ValueId = valueId;
        //    _metadataDb = metadataDb;
        //    _readValueIndex = Convert.ToInt32(valueId);
        //    _parent = parent;
        //}

        //public ExcelCellMetadataRecord(XmlReader xr, MetadataDatabase metadataDb, IndexEndpoint parent, RichDataIndexStore store)
        //   : base(store, RichDataEntities.ValueMetadataRecord)
        //{
        //    _metadataDb = metadataDb;
        //    var t = int.Parse(xr.GetAttribute("t"));
        //    var v = int.Parse(xr.GetAttribute("v"));
        //    var type = metadataDb.MetadataTypes[t - 1];
        //    TypeId = type.Id;
        //    AddRelationTo(type);
        //    var fmt = type.GetFirstOutgoingRelByType<FutureMetadataBase>();
        //    if (fmt != null)
        //    {
        //        var bk = fmt.Blocks[v];
        //        ValueId = bk.Id;
        //        AddRelationTo(bk);
        //    }
        //}

        public ExcelCellMetadataRecord(MetadataDatabase metadataDb, IndexEndpoint parent, uint typeId, uint valueId, RichDataIndexStore store)
            : base(store, RichDataEntities.CellMetadataRecord)
        {
            TypeId = typeId;
            ValueId = valueId;
            _metadataDb = metadataDb;
            //_readValueIndex = Convert.ToInt32(valueId);
            _parent = parent;
        }

        public ExcelCellMetadataRecord(XmlReader xr, MetadataDatabase metadataDb, IndexEndpoint parent, RichDataIndexStore store)
            : base(store, RichDataEntities.CellMetadataRecord)
        {
            _metadataDb = metadataDb;
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
        private readonly int _readValueIndex;



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
                var ix = _metadataDb.MetadataTypes.GetIndexById(TypeId);
                return ix.Value + 1;
            }
        }

        public int? FutureMetadataBlockIndex
        {
            get
            {
                //var bk = _metadataDb.FutureMetadata.Get(ValueId);
                var fmType = _metadataDb.MetadataTypes.Get(TypeId);
                var fm = fmType.GetFirstOutgoingRelByType<FutureMetadataBase>();
                if (fm == null) return null;
                if(fm is FutureMetadataDynamicArray futureMetadataDynamicArray)
                {
                    return futureMetadataDynamicArray.Blocks.Any(x => x.Id == ValueId) ? 0 : null;
                }
                else if(fm is FutureMetadataRichValue fmrv)
                {
                    var bk = _metadataDb.FutureMetadataRichValueBlocks.Get(ValueId);
                    var rvMetadata = bk.GetFirstIncomingRelByType<FutureMetadataBase>();
                    return rvMetadata.Blocks.GetZeroBasedIndex(ValueId);
                }
                return null;
            }
        }

        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            base.DeleteMe(relDeletions);
            var parent = _parent as ExcelCellMetadataBlock;
            if (parent != null)
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