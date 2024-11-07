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
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a bk-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelValueMetadataBlock : IndexEndpointWithSubRelations
    {
        public ExcelValueMetadataBlock(ExcelMetadata metadata, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataBlock)
        {
            _metadata = metadata;
            _store = store;
            // A value metadata block can have more than one relation to metadata types via its records
            CreateSubRelation(RichDataEntities.MetadataType);
            // A value metadata block can have more than one relation to future metadata blocks via its records
            CreateSubRelation(RichDataEntities.RichValue);
        }

        public ExcelValueMetadataBlock(XmlReader xr, ExcelMetadata metadata, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataBlock)
        {
            _metadata = metadata;
            _store = store;
            // A value metadata block can have more than one relation to metadata types via its records
            CreateSubRelation(RichDataEntities.MetadataType);
            // A value metadata block can have more than one relation to future metadata blocks via its records
            CreateSubRelation(RichDataEntities.RichValue);
            uint currentIndex = 0;
            while (xr.IsEndElementWithName("bk") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("rc"))
                {
                    var t = int.Parse(xr.GetAttribute("t"));
                    var v = int.Parse(xr.GetAttribute("v"));
                    var type = _metadata.MetadataTypes[t - 1];
                    var fmt = type.GetFirstOutgoingRelByType<FutureMetadataBase>();
                    if(fmt != null)
                    {
                        var bk = fmt.Blocks[v];
                        AddRecord(type.Id, bk.Id);
                    }
                }
                xr.Read();
                //_metadata.OnValueMetadataRead(Id, currentIndex + 1);
                currentIndex++;
            }
        }

        private readonly ExcelMetadata _metadata;
        private readonly RichDataIndexStore _store;

        public void AddRecord(uint typeId, uint valueId)
        {
            var record = new ExcelValueMetadataRecord(_metadata, this, typeId, valueId, _store);
            _metadata.ValueMetadataRecords.Add(record);
            var type = _metadata.MetadataTypes.Get(typeId);
            var typeRel = record.AddRelationTo(type, IndexType.OneBasedPointer);
            AddSubRelation(typeRel, RichDataEntities.MetadataType);
            var fm = type.GetFirstOutgoingRelByType<FutureMetadataBase>();
            if(fm != null)
            {
                var bk = fm.Blocks.Get(valueId);
                var valueRel = record.AddRelationTo(bk);
                AddSubRelation(valueRel, RichDataEntities.RichValue);
            }
        }

        public IEnumerable<ExcelValueMetadataRecord> Records
        {
            get
            {
                var result = new List<ExcelValueMetadataRecord>();
                var valuesRelation = GetSubRelations(RichDataEntities.RichValue);
                foreach(var relation in valuesRelation.SubRelations)
                {
                    var item = relation.From as ExcelValueMetadataRecord;
                    if(item != null && !item.Deleted)
                    {
                        result.Add(item);
                    }
                }
                return result;
            }
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            base.OnConnectedEntityDeleted(e);
            var valuesRelation = GetSubRelations(RichDataEntities.RichValue);
            if (e.DeletedEntity.EntityType == RichDataEntities.FutureMetadataRichDataBlock)
            {
                var relToDelete = valuesRelation.SubRelations.FirstOrDefault(x => x.To.Id == e.DeletedEntity.Id);
                if(relToDelete != null)
                {
                    var record = relToDelete.From as ExcelValueMetadataRecord;
                    // Delete the record that is connected to the deleted entity
                    relToDelete.From.DeleteMe(e.RelationDeletions);
                }
            }
            if(valuesRelation.SubRelations.Count == 0)
            {
                DeleteMe(e.RelationDeletions);
                //_metadata.OnValueMetadataBlockDeleted(Id);
            }
        }

        public void OnRecordDeleted(ExcelValueMetadataRecord record, RelationDeletions relDeletions)
        {
            if(Records.Count() <=1)
            {
                DeleteMe(relDeletions);
            }
        }
    }
}