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
using OfficeOpenXml.EventArguments;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
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
            _indexStore = _workbook.IndexStore;
        }

        private readonly ExcelWorksheet _sheet;
        private readonly ExcelWorkbook _workbook;
        private readonly CellStore<MetaDataReference> _metadataStore;
        private readonly ExcelMetadata _metadata;
        private readonly RichDataIndexStore _indexStore;

        internal bool HasRichData(int row, int col)
        {
            return HasRichData(row, col, out MetaDataReference mdr);
        }

        internal bool HasRichData(int row, int col, out uint vmId)
        {
            var result = HasRichData(row, col, out MetaDataReference mdr);
            vmId = mdr.vm;
            return result;
        }

        internal bool HasRichData(int row, int col, out MetaDataReference mdr)
        {
            mdr = _metadataStore.GetValue(row, col);
            var valueMetadataIx = mdr.vm;
            if (valueMetadataIx == 0) return false;
            return _metadata.IsRichData(mdr.vm, out uint? richDataId);
        }

        internal ExcelRichValue GetRichValueByOneBasedIndex(int index)
        {
            var id = _workbook.Metadata.Db.ValueMetadata.GetIdByIndex(index - 1);
            var rv = GetRichValue(id);
            return rv;
        }

        internal ExcelRichValue GetRichValueByOneBasedIndex(uint index)
        {
            return GetRichValueByOneBasedIndex((int)index);
        }

        internal bool HasValueBeenDeleted(uint vmId)
        {
            var valueMetaData = _metadata.Db.ValueMetadata.Get(vmId);
            return valueMetaData.Deleted;
        }

        /// <summary>
        /// Gets a rich value by its value metadata index
        /// </summary>
        /// <param name="vmId">Id of the requested <see cref="ExcelValueMetadataBlock"/></param>
        /// <returns>An <see cref="ExcelRichValue"/> instance corresponding to <paramref name="vmId"/></returns>
        internal ExcelRichValue GetRichValue(uint vmId)
        {
            var valueMetaData = _metadata.Db.ValueMetadata.Get(vmId);
            if(valueMetaData.Records == null || !valueMetaData.Records.Any()) return null;
            var valueRecord = valueMetaData.Records.First();
            var type = valueRecord.GetFirstOutgoingRelByType<ExcelMetadataType>();
            if (type == null || type.Name != FutureMetadataBase.RICHDATA_NAME) return null;
            var bk = valueRecord.GetFirstOutgoingRelByType<FutureMetadataBlock>();
            if (bk == null) return null;
            return bk.GetFirstOutgoingRelByType<ExcelRichValue>();
        }

        internal ExcelRichValue GetRichValue(int row, int col, params string[] structureTypesFilter)
        {
            return GetRichValue(row, col, out uint vmId, structureTypesFilter);
        }

        internal ExcelRichValue GetRichValue(int row, int col, out uint vmId, params string[] structureTypesFilter)
        {
            var hasRd = HasRichData(row, col, out uint vId);
            vmId = vId;
            if (!hasRd)
            {
                return null;
            }
            return GetRichValue(vmId, structureTypesFilter);
        }

        internal ExcelRichValue GetRichValue(uint vmId, params string[] structureTypesFilter)
        {
            var valueMetaData = _metadata.Db.ValueMetadata.Get(vmId);
            var bk = valueMetaData.Records[0].GetFirstOutgoingRelByType<FutureMetadataBlock>();
            var rdv = bk.GetFirstOutgoingRelByType<ExcelRichValue>();
            if (structureTypesFilter != null
                && structureTypesFilter.Any()
                && !structureTypesFilter.Contains(rdv.Structure.Type))
            {
                return null;
            }
            return rdv;
        }

        internal void AddRichData(int row, int col, ExcelRichValue richValue)
        {
            AddRichData(row, col, richValue, out uint vmId);
        }

        internal void AddRichData(int row, int col, ExcelRichValue richValue, out uint vmId)
        {
            _workbook.RichData.Db.Values.Add(richValue);

            // update the metadata
            _metadata.CreateRichValueMetadata(_workbook.RichData, richValue, out uint vId);
            vmId = vId;
        }

        /// <summary>
        /// Overwrites an existing rich data
        /// </summary>
        /// <param name="row">Row where rich data should be updated</param>
        /// <param name="col">Column where rich data should be updated</param>
        /// <param name="richValue">The new rich data that will replace the existing</param>
        /// <param name="vmId">Value metadata id</param>
        internal void UpdateRichData(int row, int col, ExcelRichValue richValue, out uint vmId)
        {
            var existingValue = GetRichValue(row, col);
            if(existingValue != null)
            {
                existingValue.DeleteMe();
            }
            AddRichData(row, col, richValue, out uint vId);
            vmId = vId;
        }

        /// <summary>
        /// Removes rich data from a cell, including removal of the vm-attribute in the worksheet cells.
        /// </summary>
        /// <param name="row">Row of the removed rich data</param>
        /// <param name="col">Column of the removed rich data</param>
        internal void DeleteRichData(int row, int col)
        {
            var existingValue = GetRichValue(row, col);
            if(existingValue != null)
            {
                existingValue.DeleteMe();
            }
            var md = _sheet._metadataStore.GetValue(row, col);
            md.vm = 0;
            _sheet._metadataStore.SetValue(row, col, md);
            _sheet.Cells[row, col].Value = null;
        }

        internal RichValueRel GetRelation(Uri target, string type)
        {
            return _workbook.RichData.GetRelation(target.OriginalString, type);
        }
    }
}
