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
    /// Table sort layer
    /// </summary>
    public class TableSortLayer : SortLayerBase
    {
        internal TableSortLayer(TableSortOptions options)
            : base(options)
        {
            _options = options;
        }

        private readonly TableSortOptions _options;

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="column"/> (zerobased) with ascending sort direction
        /// </summary>
        /// <param name="column">The column to sort</param>
        /// <returns>A <see cref="TableSortLayerBuilder"/> for adding more sort criterias</returns>
        public TableSortLayerBuilder Column(int column)
        {
            SetColumn(column);
            return new TableSortLayerBuilder(_options, this);
        }

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="column"/> (zerobased) using the supplied sort direction.
        /// </summary>
        /// <param name="column">The column to sort</param>
        /// <param name="sortOrder">Ascending or Descending sort</param>
        /// <returns>A <see cref="TableSortLayerBuilder"/> for adding more sort criterias</returns>
        public TableSortLayerBuilder Column(int column, eSortOrder sortOrder)
        {
            SetColumn(column, sortOrder);
            return new TableSortLayerBuilder(_options, this);
        }

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="columnName"/> ith ascending sort direction
        /// </summary>
        /// <param name="columnName">The name of the column to sort, see <see cref="OfficeOpenXml.Table.ExcelTableColumn.Name"/>.</param>
        /// <returns>A <see cref="TableSortLayerBuilder"/> for adding more sort criterias</returns>
        public TableSortLayerBuilder ColumnNamed(string columnName)
        {
            var ix = _options.GetColumnNameIndex(columnName);
            SetColumn(ix);
            return new TableSortLayerBuilder(_options, this);
        }

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="columnName"/> using the supplied sort direction.
        /// </summary>
        /// <param name="columnName">Name of the column to sort, see <see cref="OfficeOpenXml.Table.ExcelTableColumn.Name"/></param>
        /// <param name="sortOrder">Ascending or Descending sort</param>
        /// <returns>A <see cref="TableSortLayerBuilder"/> for adding more sort criterias</returns>
        public TableSortLayerBuilder ColumnNamed(string columnName, eSortOrder sortOrder)
        {
            var ix = _options.GetColumnNameIndex(columnName);
            SetColumn(ix, sortOrder);
            return new TableSortLayerBuilder(_options, this);
        }
    }
}
