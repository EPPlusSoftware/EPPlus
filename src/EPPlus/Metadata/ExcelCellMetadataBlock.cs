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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a bk-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelCellMetadataBlock : IndexEndpoint
    {

        public ExcelCellMetadataBlock(MetadataDatabase metadataDb)
            : base(metadataDb.IndexStore, RichDataEntities.CellMetadataBlock)
        {
            _metadataDb = metadataDb;
            _store = metadataDb.IndexStore;
            _records = new IndexedSubsetCollection<ExcelCellMetadataRecord>(metadataDb.CellMetadataRecords);
        }

        public ExcelCellMetadataBlock(XmlReader xr, MetadataDatabase metadataDb)
            : base(metadataDb.IndexStore, RichDataEntities.CellMetadataBlock)
        {
            _metadataDb = metadataDb;
            _store = metadataDb.IndexStore;
            _records = new IndexedSubsetCollection<ExcelCellMetadataRecord>(metadataDb.CellMetadataRecords);
            uint currentIndex = 0;
            while (xr.IsEndElementWithName("bk") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("rc"))
                {
                    var record = new ExcelCellMetadataRecord(xr, _metadataDb, this, _store);
                    record.AddRelationTo(this);
                    _records.Add(record);
                }
                xr.Read();
                currentIndex++;
            }
        }

        private readonly MetadataDatabase _metadataDb;
        private readonly RichDataIndexStore _store;
        private readonly IndexedSubsetCollection<ExcelCellMetadataRecord> _records;

        public void AddRecord(uint typeId, uint valueId)
        {
            var record = new ExcelCellMetadataRecord(_metadataDb, this, typeId, valueId, _store);
            _metadataDb.CellMetadataRecords.Add(record);
            var type = _metadataDb.MetadataTypes.Get(typeId);
            record.AddRelationTo(type, IndexType.OneBasedPointer);
            record.AddRelationTo(this);
            var fm = type.GetFirstOutgoingRelByType<FutureMetadataBase>();
            if (fm != null)
            {
                var bk = fm.Blocks.Get(valueId);
                record.AddRelationTo(bk);
            }
            _records.Add(record);
        }

        public IndexedSubsetCollection<ExcelCellMetadataRecord> Records => _records;

        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            // When the last record is deleted it will trigger a call to OnRecordDeleted in this class
            // which will call DeleteMe().
            base.DeleteMe(relDeletions);
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
           if(e.DeletedEntity.EntityType == RichDataEntities.CellMetadataRecord)
            {
                if(_records.Count <= 1)
                {
                    DeleteMe(e.RelationDeletions);
                }
            }
        }
    }
}