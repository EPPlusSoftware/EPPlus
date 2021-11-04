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
    public class RangeLeftToRightSortLayer : SortLayerBase
    {
        internal RangeLeftToRightSortLayer(RangeSortOptions options)
            : base(options)
        {
            options.LeftToRight = true;
            _options = options;
        }

        private readonly RangeSortOptions _options;

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="row"/> (zerobased) with ascending sort direction
        /// </summary>
        /// <param name="row">The row to sort on</param>
        /// <returns>A <see cref="RangeLeftToRightSortLayerBuilder"/> for adding more sort criterias</returns>
        public virtual RangeLeftToRightSortLayerBuilder Row(int row)
        {
            SetColumn(row);
            return new RangeLeftToRightSortLayerBuilder(_options, this);
        }

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="row"/> (zerobased) using the supplied sort direction.
        /// </summary>
        /// <param name="row">The column to sort on</param>
        /// <param name="sortOrder">Ascending or Descending sort</param>
        /// <returns>A <see cref="RangeLeftToRightSortLayerBuilder"/> for adding more sort criterias</returns>
        public RangeLeftToRightSortLayerBuilder Row(int row, eSortOrder sortOrder)
        {
            SetColumn(row, sortOrder);
            return new RangeLeftToRightSortLayerBuilder(_options, this);
        }
    }
}
