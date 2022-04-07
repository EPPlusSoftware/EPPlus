/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting
{
    /// <summary>
    /// Base class for sort layers
    /// </summary>
    public abstract class SortLayerBase
    {
        internal SortLayerBase(SortOptionsBase options)
        {
            _options = options;
        }

        private readonly SortOptionsBase _options;
        private int _column = -1;
        private int _row = -1;

        /// <summary>
        /// Sets the column
        /// </summary>
        /// <param name="column"></param>
        protected void SetColumn(int column)
        {
            _column = column;
            _options.ColumnIndexes.Add(column);
            _options.Descending.Add(false);
        }

        /// <summary>
        /// Sets the column
        /// </summary>
        /// <param name="column">Column to sort</param>
        /// <param name="sortOrder">Sort order</param>
        protected void SetColumn(int column, eSortOrder sortOrder)
        {
            _column = column;
            _options.ColumnIndexes.Add(column);
            _options.Descending.Add((sortOrder == eSortOrder.Descending));
        }

        /// <summary>
        /// Sets the row
        /// </summary>
        /// <param name="row"></param>
        protected void SetRow(int row)
        {
            _row = row;
            _options.RowIndexes.Add(row);
            _options.Descending.Add(false);
        }

        /// <summary>
        /// Sets the row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="sortOrder"></param>
        protected void SetRow(int row, eSortOrder sortOrder)
        {
            _row = row;
            _options.RowIndexes.Add(row);
            _options.Descending.Add((sortOrder == eSortOrder.Descending));
        }

        internal void SetCustomList(params string[] values)
        {
            if(_options.CustomLists.ContainsKey(_column))
            {
                throw new ArgumentException("Custom list is already defined for column index " + _column);
            }
            _options.CustomLists[_column] = values;
        }
    }
}
