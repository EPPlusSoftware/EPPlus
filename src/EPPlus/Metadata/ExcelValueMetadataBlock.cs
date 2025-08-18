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
using OfficeOpenXml.Utils.XML;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a bk-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelValueMetadataBlock : IndexEndpointReferenceCounter
    {
        public ExcelValueMetadataBlock(MetadataDatabase metadataDb)
            : base(metadataDb.IndexStore, RichDataEntities.ValueMetadataBlock)
        {
            _metadataDb = metadataDb;
            _store = metadataDb.IndexStore;
            _records = new IndexedSubsetCollection<ExcelValueMetadataRecord>(metadataDb.ValueMetadataRecords);
        }

        public ExcelValueMetadataBlock(XmlReader xr, MetadataDatabase metadataDb)
            : base(metadataDb.IndexStore, RichDataEntities.ValueMetadataBlock)
        {
            _metadataDb = metadataDb;
            _store = metadataDb.IndexStore;
            _records = new IndexedSubsetCollection<ExcelValueMetadataRecord>(metadataDb.ValueMetadataRecords);
            uint currentIndex = 0;
            while (xr.IsEndElementWithName("bk") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("rc"))
                {
                    var record = new ExcelValueMetadataRecord(xr, _metadataDb, this, _store);
                    _records.Add(record);
                }
                xr.Read();
                currentIndex++;
            }
        }

        private readonly MetadataDatabase _metadataDb;
        private readonly RichDataIndexStore _store;
        private readonly IndexedSubsetCollection<ExcelValueMetadataRecord> _records;

        public void AddRecord(uint typeId, uint valueId)
        {
            var record = new ExcelValueMetadataRecord(_metadataDb, this, typeId, valueId, _store);
            record.AddRelationTo(this);
            _metadataDb.ValueMetadataRecords.Add(record);
            var type = _metadataDb.MetadataTypes.Get(typeId);

            var fm = type.GetFirstOutgoingRelByType<FutureMetadataBase>();
            if(fm != null)
            {
                var bk = fm.Blocks.Get(valueId);
                record.AddRelationTo(bk);
            }
            _records.Add(record);
        }

        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            // When the last record is deleted it will trigger a call to OnRecordDeleted in this class
            // which will call DeleteMe().
            base.DeleteMe(relDeletions);
        }

        public IndexedSubsetCollection<ExcelValueMetadataRecord> Records => _records;

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            if (e.DeletedEntity.EntityType == RichDataEntities.ValueMetadataRecord)
            {
                if (_records.Count <= 1)
                {
                    DeleteMe(e.RelationDeletions);
                }
            }
        }
    }
}