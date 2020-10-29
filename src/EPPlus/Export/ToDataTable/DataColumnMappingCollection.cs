/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    public class DataColumnMappingCollection : List<DataColumnMapping>
    {
        private readonly Dictionary<int, DataColumnMapping> _mappingIndexes = new Dictionary<int, DataColumnMapping>();
        internal void Validate()
        {
            foreach(var mapping in this)
            {
                mapping.Validate();
            }
        }

        public void Add(int zeroBasedIndexInRange, DataColumn dataColumn)
        {
            Add(zeroBasedIndexInRange, dataColumn, null);
        }

        public void Add(int zeroBasedIndexInRange, DataColumn dataColumn, Func<object, object> transformCellValueFunc)
        {
            var mapping = new DataColumnMapping(dataColumn)
            {
                ZeroBasedColumnIndexInRange = zeroBasedIndexInRange,
                TransformCellValue = transformCellValueFunc
            };
            _mappingIndexes[mapping.ZeroBasedColumnIndexInRange] = mapping;
            Add(mapping);
            Sort((x, y) => x.ZeroBasedColumnIndexInRange.CompareTo(y.ZeroBasedColumnIndexInRange));
        }

        public void Add(int zeroBasedIndexInRange, string columnName)
        {
            Add(zeroBasedIndexInRange, columnName, null, true, null);
        }

        public void Add(int zeroBasedIndexInRange, string columnName, bool allowNull)
        {
            Add(zeroBasedIndexInRange, columnName, null, allowNull, null);
        }

        public void Add(int zeroBasedIndexInRange, string columnName, Func<object, object> transformCellValueFunc)
        {
            Add(zeroBasedIndexInRange, columnName, null, true, transformCellValueFunc);
        }

        public void Add(int zeroBasedIndexInRange, string columnName, Type columnDataType)
        {
            Add(zeroBasedIndexInRange, columnName, columnDataType, true, null);
        }

        public void Add(int zeroBasedIndexInRange, string columnName, Type columnDataType, bool allowNull)
        {
            Add(zeroBasedIndexInRange, columnName, columnDataType, allowNull, null);
        }

        public void Add(int zeroBasedIndexInRange, string columnName, Type columnDataType, bool allowNull, Func<object, object> transformCellValueFunc)
        {
            var mapping = new DataColumnMapping
            {
                ZeroBasedColumnIndexInRange = zeroBasedIndexInRange,
                DataColumnName = columnName,
                ColumnDataType = columnDataType,
                AllowNull = allowNull,
                TransformCellValue = transformCellValueFunc
            };
            mapping.Validate();
            if (this.Any(x => x.ZeroBasedColumnIndexInRange == zeroBasedIndexInRange)) throw new InvalidOperationException("Duplicate index in range: " + zeroBasedIndexInRange);
            _mappingIndexes[mapping.ZeroBasedColumnIndexInRange] = mapping;
            Add(mapping);
            Sort((x, y) => x.ZeroBasedColumnIndexInRange.CompareTo(y.ZeroBasedColumnIndexInRange));
        }

        internal DataColumnMapping GetByRangeIndex(int index)
        {
            if (!_mappingIndexes.ContainsKey(index)) return null;
            return _mappingIndexes[index];
        }

        internal bool ContainsMapping(int index)
        {
            return _mappingIndexes.ContainsKey(index);
        }
    }
}
