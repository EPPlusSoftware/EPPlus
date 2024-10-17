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
using OfficeOpenXml.Metadata.FutureMetadata;
using OfficeOpenXml.RichData.IndexRelations;
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
            return HasRichData(row, col, out MetaDataReference mdr);
        }

        internal bool HasRichData(int row, int col, out int vm)
        {
            var result = HasRichData(row, col, out MetaDataReference mdr);
            vm = mdr.vm;
            return result;
        }

        internal bool HasRichData(int row, int col, out MetaDataReference mdr)
        {
            mdr = _metadataStore.GetValue(row, col);
            var valueMetadataIx = mdr.vm;
            if (valueMetadataIx == 0) return false;
            return _metadata.IsRichData(mdr.vm);
            //if (valueMetadataIx == 0 || !_metadata.IsRichData(valueMetadataIx)) return false;
            //var vm = valueMetadataIx;
            //// vm is a 1-based index pointer
            //var vmIx = vm - 1;
            //var valueMd = _metadata.ValueMetadata[vmIx];
            //var valueRecord = valueMd.Records.First();
            //var richValue = valueRecord.GetFirstTargetByType<ExcelRichValue>();
            //var type = valueRecord.GetFirstTargetByType<ExcelMetadataType>();
            ////var type = _metadata.MetadataTypes[valueRecord.TypeIndex - 1];
            //var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            //var id = _metadata.FutureMetadata[type.Name].Blocks[valueRecord.ValueIndex].FirstTargetId;
            //if (id.HasValue == false) return false;
            //return true;
        }

        /// <summary>
        /// Gets a rich value by its value metadata index
        /// </summary>
        /// <param name="vm">1 based index</param>
        /// <returns>An <see cref="ExcelRichValue"/> instance corresponding to <paramref name="vm"/></returns>
        internal ExcelRichValue GetRichValue(int vm)
        {
            var valueMetaData = _metadata.ValueMetadata[vm - 1];
            var valueRecord = valueMetaData.Records[0];
            var type = valueRecord.GetFirstTargetByType<ExcelMetadataType>();
            //var type = _metadata.MetadataTypes[valueRecord.TypeIndex - 1];
            if (type == null || type.Name != FutureMetadataBase.RICHDATA_NAME) return null;
            //var fmd = _metadata.FutureMetadata[type.Name];
            var bk = valueRecord.GetFirstTargetByType<FutureMetadataRichValueBlock>();
            //var id = fmd.Blocks[valueRecord.ValueIndex].FirstTargetId;
            if (bk == null) return null;
            return bk.GetFirstTargetByType<ExcelRichValue>();
            //return _workbook.RichData.Values.GetItem(id.Value);
        }

        private ExcelRichValue GetRichValue(int row, int col, out int? richValueIndex)
        {
            var result = GetRichValue(row, col, out int? rvIndex, null);
            richValueIndex = rvIndex;
            return result;
        }

        internal ExcelRichValue GetRichValue(int row, int col, params string[] structuretypesFilter)
        {
            return GetRichValue(row, col, out int? rvIx, structuretypesFilter);
        }

        internal ExcelRichValue GetRichValue(int row, int col, out int? richValueIndex, params string[] structureTypesFilter)
        {
            richValueIndex = null;
            if (!HasRichData(row, col, out int vm)) return null;
            var valueMetaData = _metadata.ValueMetadata[vm - 1];
            var valueRecord = valueMetaData.Records[0];
            // var type = _metadata.MetadataTypes[valueRecord.TypeIndex - 1];
            var bk = valueRecord.GetFirstTargetByType<FutureMetadataRichValueBlock>();
            //var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            var rdv = bk.GetFirstTargetByType<ExcelRichValue>();
            //var rdv = _workbook.RichData.Values.Items[valueRecord.ValueIndex];
            if(structureTypesFilter != null 
                && structureTypesFilter.Any()
                && !structureTypesFilter.Contains(rdv.Structure.Type))
            {
                return null;
            }
            richValueIndex = valueRecord.ValueIndex;
            return rdv;
        }

        internal ExcelRichValueStructure GetStructure(RichDataStructureTypes structureType)
        {
            //var structureId = _workbook.RichData.Structures.GetStructureId(structureType);
            //return _workbook.RichData.Structures.StructureItems[structureId];
            return _workbook.RichData.Structures.GetByType(structureType);
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

        internal void AddRichData(int row, int col, ExcelRichValue richValue)
        {
            var rvIx = _workbook.RichData.Values.Count;
            _workbook.RichData.Values.Add(richValue);

            // update the metadata
            _metadata.CreateRichValueMetadata(_workbook.RichData, rvIx, out int vm);
            var md = _sheet._metadataStore.GetValue(row, col);
            md.vm = vm;
            _sheet._metadataStore.SetValue(row, col, md);
        }

        /// <summary>
        /// Overwrites an existing rich data
        /// </summary>
        /// <param name="richValueIndex">Index of the rich data to update in the _workbook.RichData.Values.Items collection</param>
        /// <param name="richValue">The new rich data that will replace the existing</param>
        /// <param name="targetUri"></param>
        internal void UpdateRichData(int row, int col, ExcelRichValue richValue)
        {
            var existingValue = GetRichValue(row, col, out int? rvIndex);
            if(!rvIndex.HasValue)
            {
                AddRichData(row, col, richValue);
            }
            var richValueIndex = rvIndex.Value;
            //var existingValue = _workbook.RichData.Values.Items[richValueIndex];

            //if(existingValue.StructureId == richValue.StructureId && targetUri != null)
            //{
            //    // at this stage we only support one relation per rich value
            //    var existingRelation = existingValue.Structure.GetFirstRelationIndex();
            //    if (existingRelation.HasValue && targetUri != null)
            //    {
            //        _workbook.RichData.RichValueRels.SetNewTarget(existingRelation.Value, targetUri);
            //    }
            //}
            foreach(var key in richValue.Structure.Keys)
            {
                if(key.IsRelation)
                {
                    var targetUri = richValue.GetRelation(key.Name, out int? relIx);
                    if(targetUri != null && relIx.HasValue)
                    {
                        _workbook.RichData.RichValueRels.SetNewTarget(relIx.Value, targetUri);
                    }

                }
            }
            _workbook.RichData.Values[richValueIndex] = richValue;
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
            var bk = valueRecord.GetFirstTargetByType<FutureMetadataRichValueBlock>();
            if (bk != null) return false;
            var rv = bk.GetFirstTargetByType<ExcelRichValue>();
            rv.DeleteMe();
            return true;
            //var type = _metadata.MetadataTypes[valueRecord.TypeIndex - 1];
            //var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            //var ix = _metadata.FutureMetadata[type.Name].Blocks[valueRecord.ValueIndex].FirstTargetIndex;
            //if (!ix.HasValue) return false;
            //return _workbook.RichData.Deletions.DeleteRichData(vmIx, ix.Value);
        }

    }
}
