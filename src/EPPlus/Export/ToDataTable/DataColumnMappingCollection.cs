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
    /// <summary>
    /// A collection of <see cref="DataColumnMapping"/>s that will be used when reading data from the source range.
    /// </summary>
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

        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="dataColumn">The destination <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        public void Add(int zeroBasedIndexInRange, DataColumn dataColumn)
        {
            Add(zeroBasedIndexInRange, dataColumn, null);
        }

        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="dataColumn">The destination <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        /// <param name="transformCellValueFunc">A function that casts/transforms the value before it is written to the <see cref="DataTable"/></param>
        /// <seealso cref="DataColumnMapping.TransformCellValue"/>
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
        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="columnName">Name of the <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        public void Add(int zeroBasedIndexInRange, string columnName)
        {
            Add(zeroBasedIndexInRange, columnName, null, true, null);
        }

        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="columnName">Name of the <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        /// <param name="allowNull">Indicates if values read from the source range can be null</param>
        public void Add(int zeroBasedIndexInRange, string columnName, bool allowNull)
        {
            Add(zeroBasedIndexInRange, columnName, null, allowNull, null);
        }

        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="columnName">Name of the <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        /// <param name="transformCellValueFunc">A function that casts/transforms the value before it is written to the <see cref="DataTable"/></param>
        public void Add(int zeroBasedIndexInRange, string columnName, Func<object, object> transformCellValueFunc)
        {
            Add(zeroBasedIndexInRange, columnName, null, true, transformCellValueFunc);
        }

        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="columnName">Name of the <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        /// <param name="columnDataType"><see cref="Type"/> of the <see cref="DataColumn"/></param>
        public void Add(int zeroBasedIndexInRange, string columnName, Type columnDataType)
        {
            Add(zeroBasedIndexInRange, columnName, columnDataType, true, null);
        }

        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="columnName">Name of the <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        /// <param name="columnDataType"><see cref="Type"/> of the <see cref="DataColumn"/></param>
        /// <param name="allowNull">Indicates if values read from the source range can be null</param>
        public void Add(int zeroBasedIndexInRange, string columnName, Type columnDataType, bool allowNull)
        {
            Add(zeroBasedIndexInRange, columnName, columnDataType, allowNull, null);
        }

        /// <summary>
        /// Adds a <see cref="DataColumnMapping"/>
        /// </summary>
        /// <param name="zeroBasedIndexInRange">Zero based index of the column in the source range</param>
        /// <param name="columnName">Name of the <see cref="DataColumn"/> in the <see cref="DataTable"/></param>
        /// <param name="columnDataType"><see cref="Type"/> of the <see cref="DataColumn"/></param>
        /// <param name="allowNull">Indicates if values read from the source range can be null</param>
        /// <param name="transformCellValueFunc">A function that casts/transforms the value before it is written to the <see cref="DataTable"/></param>
        /// <seealso cref="DataColumnMapping.TransformCellValue"/>
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

        internal bool SetNewName(string existingName, string newName)
        {
            for(var i = 0; i < Count; i++)
            {
                var mapping = this[i];
                if(mapping.DataColumnName == existingName)
                {
                    mapping.DataColumnName = newName;
                    return true;
                }
            }
            return false;
        }
    }
}
