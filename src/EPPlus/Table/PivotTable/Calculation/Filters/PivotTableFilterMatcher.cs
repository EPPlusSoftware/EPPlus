﻿using EPPlusTest.Table.PivotTable;
using System;
using System.Collections.Generic;
using OfficeOpenXml.Table.PivotTable.Filter;
 /*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.Table.PivotTable.Calculation.Filters
{
    internal static class PivotTableFilterMatcher
    {
   //     internal static bool IsFiltered(ExcelPivotTable pivotTable, PivotTableCacheRecords recs, int r)
   //     {
   //         if (pfCount > 0)
   //         {
   //             if(IsHiddenByPageField(pivotTable, recs, r))
   //             {
   //                 return true;
   //             }
   //         }

			//if (filterCount > 0)
   //         {                
   //             if(IsHiddenByRowColumnFilter(pivotTable, recs, r))
   //             { 
   //                 return true; 
   //             }
   //         }
   //         return false;
   //     }

        /// <summary>
        /// Returns true if the record is hidden by a page filter in the pivot table
        /// </summary>
        /// <param name="pivotTable">The pivot table</param>
        /// <param name="recs">The pivot cache records</param>
        /// <param name="r">The record index</param>
        /// <returns></returns>
        internal static bool IsHiddenByPageField(ExcelPivotTable pivotTable, PivotTableCacheRecords recs, int r)
        {
            foreach (var p in pivotTable.PageFields)
            {
                if (p.MultipleItemSelectionAllowed == false)
                {
                    var ix = p.PageFieldSettings.SelectedItem;
                    if(ix < 0 || ix > p.Items.Count || ix.Equals(recs.CacheItems[p.Index][r]))
                    {
                        return true;
                    }
                }
                else
                {
                    var itemIx = recs.CacheItems[p.Index][r];
                    if(!p.Items.HiddenItemIndex.Exists(x => x.Equals(itemIx)))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        /// <summary>
        /// Returns true if a record is hidden by a caption/date or numeric filter
        /// </summary>
        /// <param name="pivotTable"></param>
        /// <param name="captionFilters"></param>
        /// <param name="recs"></param>
        /// <param name="r"></param>
        /// <returns></returns>
        internal static bool IsHiddenByRowColumnFilter(ExcelPivotTable pivotTable, List<ExcelPivotTableFilter> captionFilters, PivotTableCacheRecords recs, int r)
        {
            foreach (var f in captionFilters)
            {
				var fld = pivotTable.Fields[f.Fld];
				if (fld.IsColumnField || fld.IsRowField)
                {
                    if (f.MatchesLabel(pivotTable, recs, r))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

    }
}