/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml.Table.PivotTable;
using System;

namespace OfficeOpenXml.Table.PivotTable.Filter
{
    /// <summary>
    /// A collection of pivot filters for a pivot table field
    /// </summary>
    public class ExcelPivotTableFieldFilterCollection : ExcelPivotTableFilterBaseCollection
    {
        internal ExcelPivotTableFieldFilterCollection(ExcelPivotTableField field) : base(field)
        {
        }

        /// <summary>
        /// Adds a caption (label) filter for a pivot tabel field
        /// </summary>
        /// <param name="type">The type of pivot table caption filter</param>
        /// <param name="value1">Value 1</param>
        /// <param name="value2">Value 2. Set to null, if not used</param>
        /// <returns></returns>
        public ExcelPivotTableFilter AddCaptionFilter(ePivotTableCaptionFilterType type, string value1, string value2=null)
        {
            ExcelPivotTableFilter filter = CreateFilter();
            filter.Type = (ePivotTableFilterType)type;

            filter.StringValue1 = value1;
            filter.Value1 = value1;
            filter.Value2 = value2;
            if (!string.IsNullOrEmpty(value2))
            {
                filter.StringValue2 = value2;
            }

            switch (type)
            {
                case ePivotTableCaptionFilterType.CaptionEqual:
                    filter.CreateValueFilter();
                    break;
                default:
                    filter.CreateCaptionCustomFilter(type);
                    break;
            }

            filter.Filter.Save();
            _filters.Add(filter);
            return filter;
        }
        /// <summary>
        /// Adds a date filter for a pivot table field
        /// </summary>
        /// <param name="type">The type of pivot table filter.</param>
        /// <param name="value1">Value 1</param>
        /// <param name="value2">Value 2. Set to null, if not used</param>
        /// <returns>The pivot table filter</returns>
        /// <exception cref="ArgumentNullException">Thrown if value is between and <paramref name="value2"/> is null</exception>
        public ExcelPivotTableFilter AddDateValueFilter(ePivotTableDateValueFilterType type, DateTime value1, DateTime? value2 = null)
        {
            if(value2.HasValue==false &&
               (type == ePivotTableDateValueFilterType.DateBetween || 
               type == ePivotTableDateValueFilterType.DateNotBetween))
            {
                throw new ArgumentNullException("value2", "Between filters require two values");
            }

            ExcelPivotTableFilter filter = CreateFilter();
            filter.Type = (ePivotTableFilterType)type;
            filter.Value1 = value1;
            filter.Value2 = value2;
            filter.CreateDateCustomFilter(type);

            filter.Filter.Save();
            _filters.Add(filter);
            return filter;
        }
        /// <summary>
        /// Adds a date period filter for a pivot table field.
        /// </summary>
        /// <param name="type">The type of field.</param>
        /// <returns>The pivot table filter</returns>
        public ExcelPivotTableFilter AddDatePeriodFilter(ePivotTableDatePeriodFilterType type)
        {
            ExcelPivotTableFilter filter = CreateFilter();
            filter.Type = (ePivotTableFilterType)type;

            filter.CreateDateDynamicFilter(type);

            filter.Filter.Save();
            _filters.Add(filter);
            return filter;
        }
        /// <summary>
        /// Adds a pivot table value filter.
        /// </summary>
        /// <param name="type">The type of value filter</param>
        /// <param name="dataField">The data field to apply the filter to</param>
        /// <param name="value1">Value 1</param>
        /// <param name="value2">Value 2. Used with <see cref="ePivotTableValueFilterType.ValueBetween"/> and <see cref="ePivotTableValueFilterType.ValueNotBetween"/> </param>
        /// <returns>The pivot table filter</returns>
        /// <exception cref="ArgumentException">If the data field is not present in the pivot table.</exception>
        /// <exception cref="ArgumentNullException">If value2 is not set when type is set to between</exception>
        public ExcelPivotTableFilter AddValueFilter(ePivotTableValueFilterType type, ExcelPivotTableDataField dataField, object value1, object value2 = null)
        {
            var dfIx = _table.DataFields._list.IndexOf(dataField);
            if(dfIx<0)
            {
                throw new ArgumentException("This datafield is not in the pivot tables DataFields collection", "dataField");
            }
            return AddValueFilter(type, dfIx, value1, value2);
        }
        /// <summary>
        /// Adds a pivot table value filter.
        /// </summary>
        /// <param name="type">The type of value filter</param>
        /// <param name="dataFieldIndex">The index of the <see cref="ExcelPivotTableDataField"/> to apply the filter to.</param>
        /// <param name="value1">Value 1</param>
        /// <param name="value2">Value 2. Used with <see cref="ePivotTableValueFilterType.ValueBetween"/> and <see cref="ePivotTableValueFilterType.ValueNotBetween"/></param>
        /// <returns>The pivot table filter</returns>
        /// <exception cref="ArgumentException">If the data field is not present in the pivot table.</exception>
        /// <exception cref="ArgumentNullException">If value2 is not set when type is set to between</exception>
        public ExcelPivotTableFilter AddValueFilter(ePivotTableValueFilterType type, int dataFieldIndex, object value1, object value2 = null)
        {
            if(dataFieldIndex<0 || dataFieldIndex >= _table.DataFields.Count)
            {
                throw new ArgumentException("dataFieldIndex must point to an item in the pivot tables DataFields collection", "dataFieldIndex");
            }

            if (value2 == null &&
               (type == ePivotTableValueFilterType.ValueBetween ||
               type == ePivotTableValueFilterType.ValueNotBetween))
            {
                throw new ArgumentNullException("value2", "Between filters require two values");
            }

            ExcelPivotTableFilter filter = CreateFilter();
            filter.Type = (ePivotTableFilterType)type;
            filter.Value1 = value1;
            filter.Value2 = value2;
            filter.MeasureFldIndex = dataFieldIndex;

            filter.CreateValueCustomFilter(type);

            filter.Filter.Save();
            _filters.Add(filter);
            return filter;
        }
        /// <summary>
        /// Adds a top 10 filter to the field
        /// </summary>
        /// <param name="type">The top-10 filter type</param>
        /// <param name="dataField">The datafield within the pivot table</param>
        /// <param name="value">The top or bottom value to relate to </param>
        /// <param name="isTop">Top or bottom. true is Top, false is Bottom</param>
        /// <returns></returns>
        public ExcelPivotTableFilter AddTop10Filter(ePivotTableTop10FilterType type, ExcelPivotTableDataField dataField, double value, bool isTop = true)
        {
            var dfIx = _table.DataFields._list.IndexOf(dataField);
            if (dfIx < 0)
            {
                throw new ArgumentException("This data field is not in the pivot tables DataFields collection", "dataField");
            }
            return AddTop10Filter(type, dfIx, value, isTop);

        }
        /// <summary>
        /// Adds a top 10 filter to the field
        /// </summary>
        /// <param name="type">The top-10 filter type</param>
        /// <param name="dataFieldIndex">The index to the data field within the pivot tables DataField collection</param>
        /// <param name="value">The top or bottom value to relate to </param>
        /// <param name="isTop">Top or bottom. true is Top, false is Bottom</param>
        /// <returns></returns>
        public ExcelPivotTableFilter AddTop10Filter(ePivotTableTop10FilterType type, int dataFieldIndex, double value, bool isTop=true)
        {
            if (dataFieldIndex < 0 || dataFieldIndex >= _table.DataFields.Count)
            {
                throw new ArgumentException("dataFieldIndex must point to an item in the pivot tables DataFields collection", "dataFieldIndex");
            }

            ExcelPivotTableFilter filter = CreateFilter();
            filter.Type = (ePivotTableFilterType)type;
            filter.Value1 = value;
            filter.MeasureFldIndex = dataFieldIndex;

            filter.CreateTop10Filter(type, isTop, value);

            filter.Filter.Save();
            _filters.Add(filter);
            return filter;
        }
    }
}
