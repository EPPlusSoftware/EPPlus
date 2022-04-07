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
    /// Used to create sort criterias for sorting a range.
    /// </summary>
    public class TableSortLayerBuilder
    {
        internal TableSortLayerBuilder(TableSortOptions options, TableSortLayer sortLayer)
        {
            _options = options;
            _sortLayer = sortLayer;
        }

        private readonly TableSortOptions _options;
        private readonly TableSortLayer _sortLayer;

        /// <summary>
        /// Add a new Sort layer.
        /// </summary>
        public TableSortLayer ThenSortBy
        {
            get
            {
                return new TableSortLayer(_options);
            }
        }

        /// <summary>
        /// Use a custom list for sorting on the current Sort layer.
        /// </summary>
        /// <param name="values">An array of strings defining the sort order</param>
        /// <returns>A <see cref="TableSortLayerBuilder"/></returns>
        public TableSortLayerBuilder UsingCustomList(params string[] values)
        {
            _sortLayer.SetCustomList(values);
            return this;
        }
    }
}
