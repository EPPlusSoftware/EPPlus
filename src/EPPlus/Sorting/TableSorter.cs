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
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting
{
    internal class TableSorter
    {
        public TableSorter(ExcelTable table)
        {
            _table = table;
        }

        private readonly ExcelTable _table;

        public void Sort(TableSortOptions options)
        {
            _table.DataRange.Sort(options, _table);
        }

        public void Sort(Action<TableSortOptions> configuration)
        {
            var options = new TableSortOptions(_table);
            configuration(options);
            Sort(options);
        }
    }
}
