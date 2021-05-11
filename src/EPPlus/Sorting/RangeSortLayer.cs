/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting
{
    /// <summary>
    /// This class represents
    /// </summary>
    public class RangeSortLayer : SortLayerBase
    {
        internal RangeSortLayer(RangeSortOptions options)
            : base(options)
        {
            _options = options;
        }

        private readonly RangeSortOptions _options;

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="column"/> (zerobased) with ascending sort direction
        /// </summary>
        /// <param name="column">The column to sort</param>
        /// <returns>A <see cref="RangeSortLayerBuilder"/> for adding more sort criterias</returns>
        public virtual RangeSortLayerBuilder Column(int column)
        {
            SetColumn(column);
            return new RangeSortLayerBuilder(_options, this);
        }

        /// <summary>
        /// Sorts by the column that corresponds to the <paramref name="column"/> (zerobased) using the supplied sort direction.
        /// </summary>
        /// <param name="column">The column to sort</param>
        /// <param name="direction">Ascending or Descending sort</param>
        /// <returns>A <see cref="RangeSortLayerBuilder"/> for adding more sort criterias</returns>
        public RangeSortLayerBuilder Column(int column, eSortDirection direction)
        {
            SetColumn(column, direction);
            return new RangeSortLayerBuilder(_options, this);
        }
    }
}
