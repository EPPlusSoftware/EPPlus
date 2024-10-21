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
        public ExcelValueMetadataBlock(ExcelMetadata metadata, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataBlock)
        {
            _metadata = metadata;
            _store = store;
            // A value metadata block can have more than one relation to metadata types via its records
            _typeRelation = new IndexRelationWithSubRelations(this, RichDataEntities.MetadataType, IndexType.SubRelations);
            store.AddRelationWithSubRelations(_typeRelation);
            // A value metadata block can have more than one relation to future metadata blocks via its records
            _valuesRelation = new IndexRelationWithSubRelations(this, RichDataEntities.RichValue, IndexType.SubRelations);
            store.AddRelationWithSubRelations(_valuesRelation);

        }

        public ExcelValueMetadataBlock(XmlReader xr, ExcelMetadata metadata, RichDataIndexStore store)
            : base(store, RichDataEntities.ValueMetadataBlock)
        {
            _metadata = metadata;
            _store = store;
            // A value metadata block can have more than one relation to metadata types via its records
            _typeRelation = new IndexRelationWithSubRelations(this, RichDataEntities.MetadataType, IndexType.SubRelations);
            // A value metadata block can have more than one relation to future metadata blocks via its records
            _valuesRelation = new IndexRelationWithSubRelations(this, RichDataEntities.RichValue, IndexType.SubRelations);
            store.AddRelationWithSubRelations(_valuesRelation);
            store.AddRelationWithSubRelations(_typeRelation);
            while (xr.IsEndElementWithName("bk") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("rc"))
                {
                    var t = int.Parse(xr.GetAttribute("t"));
                    var v = int.Parse(xr.GetAttribute("v"));
                    var type = _metadata.MetadataTypes[t - 1];
                    var fmt = type.GetFirstTargetByType<FutureMetadataBase>();
                    var bk = fmt.Blocks[v];
                    AddRecord(type.Id, bk.Id);
                    //Records.Add(new ExcelValueMetadataRecord(metadata, this, t, v, store));
                }
                xr.Read();
            }
        }

        private readonly ExcelMetadata _metadata;
        private readonly RichDataIndexStore _store;
        private readonly IndexRelationWithSubRelations _typeRelation;
        private readonly IndexRelationWithSubRelations _valuesRelation;

        public void AddRecord(uint typeId, uint valueId)
        {
            var record = new ExcelValueMetadataRecord(_metadata, this, typeId, valueId, _store);
            _metadata.ValueMetadataRecords.Add(record);
            var type = _metadata.MetadataTypes.Get(typeId);
            var typeRel = _metadata.MetadataTypes.CreateRelation(record, type, IndexType.OneBasedPointer);
            _typeRelation.SubRelations.Add(typeRel);
            var fm = type.GetFirstTargetByType<FutureMetadataBase>();
            if(fm != null)
            {
                var bk = fm.Blocks.Get(valueId);
                var valueRel = _metadata.FutureMetadataBlocks.CreateRelation(record, bk, IndexType.ZeroBasedPointer);
                _valuesRelation.SubRelations.Add(valueRel);
            }
        }

        public IEnumerable<ExcelValueMetadataRecord> Records
        {
            get
            {
                var result = new List<ExcelValueMetadataRecord>();
                foreach(var relation in _valuesRelation.SubRelations)
                {
                    var item = relation.From as ExcelValueMetadataRecord;
                    if(item != null)
                    {
                        result.Add(item);
                    }
                }
                return result;
            }
        }
    }
}