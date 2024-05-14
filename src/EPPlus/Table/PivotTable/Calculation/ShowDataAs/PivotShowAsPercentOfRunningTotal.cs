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
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfRunningTotal : PivotShowAsRunningTotal
	{
		internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems)
		{
			CalculateRunningTotal(df, fieldIndex, ref calculatedItems, true);
            if (_bf < 0)
            {
                return;
            }
            foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
            {
                if (IsSumBefore(key.Key, _bf, fieldIndex, _colFieldsStart))
                {
                    calculatedItems[key.Key] = null;
                }
                else
                {
                    if (IsSumAfter(key.Key, _bf, fieldIndex, _colFieldsStart))
                    {
                        var parentKey = GetParentKey(key.Key, _keyCol);
                        var parentValue = calculatedItems[parentKey];
						if (parentValue is double pv)
						{
							if (calculatedItems[key.Key] is double v)
							{
								calculatedItems[key.Key] = (double)v / pv;
							}
						}
                        else if (calculatedItems[key.Key] is not ExcelErrorValue)
                        {
                            calculatedItems[key.Key] = parentValue;
                        }
                    }
                    else
                    {
                        calculatedItems[key.Key] = 1D;
                    }
                }
            }
        }
    }
}
