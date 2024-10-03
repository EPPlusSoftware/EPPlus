/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData.Mappings;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.RichValues.Relations;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.RichData
{
    internal class RichDataStore
    {
        public RichDataStore(ExcelWorksheet sheet)
        {
            _sheet = sheet;
            _workbook = sheet.Workbook;
            _metadataStore = sheet._metadataStore;
            _metadata = sheet.Workbook.Metadata;
        }

        private readonly ExcelWorksheet _sheet;
        private readonly ExcelWorkbook _workbook;
        private readonly CellStore<MetaDataReference> _metadataStore;
        private readonly ExcelMetadata _metadata;

        internal bool HasRichData(int row, int col)
        {
            return HasRichData(row, col, out int vm, out ExcelMetadataRecord valueRecord);
        }

        internal bool HasRichData(int row, int col, out int vm)
        {
            return HasRichData(row, col, out vm, out ExcelMetadataRecord valueRecord);
        }

        internal bool HasRichData(int row, int col, out int vm, out ExcelMetadataRecord valueRecord)
        {
            vm = 0;
            valueRecord = null;
            var valueMetadataIx = _metadataStore.GetValue(row, col).vm;
            if (valueMetadataIx == 0 || !_metadata.IsRichData(valueMetadataIx)) return false;
            vm = valueMetadataIx;
            // vm is a 1-based index pointer
            var vmIx = vm - 1;
            var valueMd = _metadata.ValueMetadata[vmIx];
            valueRecord = valueMd.Records.First();
            var type = _metadata.MetadataTypes[valueRecord.TypeIndex - 1];
            var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            var ix = _metadata.FutureMetadata[type.Name].Types[valueRecord.ValueIndex].AsRichData.Index;
            if (_workbook.RichData.Deletions.IsDeleted(vmIx, ix)) return false;
            return true;
        }

        internal ExcelRichValue GetRichValue(int row, int col, params string[] structuretypesFilter)
        {
            return GetRichValue(row, col, out int? rvIx, structuretypesFilter);
        }

        internal ExcelRichValue GetRichValue(int row, int col, out int? richValueIndex, params string[] structureTypesFilter)
        {
            richValueIndex = null;
            if (!HasRichData(row, col, out int vm, out ExcelMetadataRecord valueRecord)) return null;
            var type = _metadata.MetadataTypes[valueRecord.TypeIndex - 1];
            var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            var rdv = _workbook.RichData.Values.Items[valueRecord.ValueIndex];
            if(structureTypesFilter != null && !structureTypesFilter.Contains(rdv.Structure.Type))
            {
                return null;
            }
            richValueIndex = valueRecord.ValueIndex;
            return rdv;
        }

        internal ExcelRichValueStructure GetStructure(RichDataStructureTypes structureType)
        {
            var structureId = _workbook.RichData.Structures.GetStructureId(structureType);
            return _workbook.RichData.Structures.StructureItems[structureId];
        }

        internal int CreateRichValueRelation(RichDataStructureTypes structureType, Uri rvRelUri)
        {
            var structure = GetStructure(structureType);
            var index = structure.GetFirstRelationIndex();
            if(!index.HasValue)
            {
                throw new InvalidOperationException($"Cannot create a relation from structure {structure.Type}/{structure.StructureType}");
            }
            var rel = structure.Keys[index.Value].Name;
            var relationshipType = RichValueRelationMappings.GetSchema(rel);
            _workbook.RichData.RichValueRels.AddItem(rvRelUri, relationshipType, out int relIx);
            return relIx;
        }

        internal void AddRichData(ExcelRichValue richValue, RichDataStructureTypes structureType, out int vmIndex)
        {
            _workbook.RichData.Values.Items.Add(richValue);

            // update the metadata
            _metadata.CreateRichValueMetadata(_workbook.RichData, out int vm);
            vmIndex = vm;
        }

        internal void UpdateRichData(int richValueIndex, ExcelRichValue richValue, Uri targetUri)
        {
            var existingValue = _workbook.RichData.Values.Items[richValueIndex];

            if(existingValue.StructureId == richValue.StructureId)
            {
                // at this stage we only support one relation per rich value
                var existingRelation = existingValue.Structure.GetFirstRelationIndex();
                if (existingRelation.HasValue && targetUri != null)
                {
                    _workbook.RichData.RichValueRels.SetNewTarget(existingRelation.Value, targetUri);
                }
            }
            _workbook.RichData.Values.Items[richValueIndex] = richValue;
        }

        internal RichValueRel GetRelation(int relationIndex)
        {
            return _workbook.RichData.GetRelation(relationIndex);
        }

        internal RichValueRel GetRelation(Uri target, string type)
        {
            return _workbook.RichData.GetRelation(target.OriginalString, type);
        }

        internal bool DeleteRichData(int row, int col)
        {
            var vm = _metadataStore.GetValue(row, col).vm;
            if (vm == 0 || !_metadata.IsRichData(vm)) return false;
            var vmIx = vm - 1;
            var valueMd = _metadata.ValueMetadata[vmIx];
            var valueRecord = valueMd.Records.First();
            var type = _metadata.MetadataTypes[valueRecord.TypeIndex - 1];
            var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            var ix = _metadata.FutureMetadata[type.Name].Types[valueRecord.ValueIndex].AsRichData.Index;
            return _workbook.RichData.Deletions.DeleteRichData(vmIx, ix);
        }

    }
}
