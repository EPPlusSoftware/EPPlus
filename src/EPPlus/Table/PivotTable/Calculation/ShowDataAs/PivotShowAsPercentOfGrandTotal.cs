/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/18/2024         EPPlus Software AB       EPPlus 7.2
*************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfGrandTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {
            var totalKey = GetKey(fieldIndex.Count);            
            var t = calculatedItems[totalKey];
            if(t is double total)
            {
                foreach(var key in calculatedItems.Index)
                {
                    if (calculatedItems[key.Key] is double d)
                    {
                        calculatedItems[key.Key] = d / total;
                    }
                }
            }
            else //Not a double, its an excel error.
            {
                foreach (var key in calculatedItems.Index)
                {
                    if (calculatedItems[key.Key] is double d)
                    {
                        calculatedItems[key.Key] = t;
                    }
                }
            }
        }
    }
}
