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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.Core.CellStore
{
    internal class MetadataCellStore : CellStore<MetaDataReference>
    {
        public MetadataCellStore(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
            _metadata = _worksheet.Workbook.Metadata;
        }

        private readonly ExcelWorksheet _worksheet;
        private readonly ExcelMetadata _metadata;

        internal override void SetValue(int row, int column, MetaDataReference value)
        {
            HandleMetadataReferences(row, column, value);
            base.SetValue(row, column, value);
        }

        internal override void Delete(int fromRow, int fromCol, int rows, int columns)
        {
            HandleMetadataReferencesRange(fromRow, fromCol, rows, columns);
            base.Delete(fromRow, fromCol, rows, columns);
        }

        internal override void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
        {
            HandleMetadataReferencesRange(fromRow, fromCol, rows, columns);
            base.Delete(fromRow, fromCol, rows, columns, shift);
        }

        internal override void Clear(int fromRow, int fromCol, int rows, int columns)
        {
            base.Clear(fromRow, fromCol, rows, columns);
        }

        private void HandleMetadataReferencesRange(int fromRow, int fromCol, int rows, int columns)
        {
            var cse = new CellStoreEnumerator<MetaDataReference>(this, fromRow, fromCol, fromRow + rows - 1, fromCol + columns - 1);
            foreach(var cell in cse)
            { 
                if (cell.cm > 0)
                {
                    var existingCmBk = _metadata.Db.CellMetadata.Get(cell.cm);
                    if (existingCmBk != null)
                    {
                        existingCmBk.DecreaseReferences();
                    }
                }
                if (cell.vm > 0)
                {
                    var existingVmBk = _metadata.Db.CellMetadata.Get(cell.vm);
                    if (existingVmBk != null)
                    {
                        existingVmBk.DecreaseReferences();
                    }
                }
            }
        }

        private void HandleMetadataReferences(int row, int column, MetaDataReference value)
        {
            var existingMetadata = GetValue(row, column);
            if (existingMetadata.cm > 0 && existingMetadata.cm != value.cm)
            {
                var existingCmBk = _metadata.Db.CellMetadata.Get(existingMetadata.cm);
                if (existingCmBk != null)
                {
                    existingCmBk.DecreaseReferences();
                }
            }
            var newCmBk = _metadata.Db.CellMetadata.Get(value.cm);
            if (newCmBk != null)
            {
                newCmBk.IncreaseReferences();
            }
            if (existingMetadata.vm > 0 && existingMetadata.vm != value.vm)
            {
                var existingVmBk = _metadata.Db.ValueMetadata.Get(existingMetadata.vm);
                if (existingVmBk != null)
                {
                    existingVmBk.DecreaseReferences();
                }
            }
            var newVmBk = _metadata.Db.ValueMetadata.Get(value.vm);
            if (newVmBk != null)
            {
                newVmBk.IncreaseReferences();
            }
        }

        internal override MetaDataReference GetValue(int Row, int Column)
        {
            return base.GetValue(Row, Column);
        }
    }
}
