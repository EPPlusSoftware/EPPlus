/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/7/2023         EPPlus Software AB       EPPlus 7.0.4
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class ColumnInfoCollection : List<ColumnInfo>
    {
        internal void ReindexAndSortColumns()
        {
            var index = 0;
            Sort((a, b) =>
            {
                //var so1 = a.SortOrderLevels;
                //var so2 = b.SortOrderLevels;
                //var maxIx = so1.Count < so2.Count ? so1.Count : so2.Count;
                //for (var ix = 0; ix < maxIx; ix++)
                //{
                //    var aVal = so1[ix];
                //    var bVal = so2[ix];
                //    if (aVal.CompareTo(bVal) == 0) continue;
                //    return aVal.CompareTo(bVal);
                //}
                //return a.Index.CompareTo(b.Index);
                var p1 = a.Path;
                var p2 = b.Path;
                var maxIx = p1.Depth < p2.Depth ? p1.Depth : p2.Depth;
                for(var ix = 0; ix < maxIx; ix++)
                {
                    var aVal = p1.Get(ix).SortOrder;
                    var bVal = p2.Get(ix).SortOrder;
                    if (aVal.CompareTo(bVal) == 0) continue;
                    return aVal.CompareTo(bVal);
                }
                return a.Index.CompareTo(b.Index);
            });
            this.ForEach(x => x.Index = index++);
        }
    }
}
