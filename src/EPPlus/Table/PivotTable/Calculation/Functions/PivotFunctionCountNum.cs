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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionCountNums : PivotFunction
    {
        internal override void AddItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
        {
            if (value is KahanSum || IsNumeric(value))
            {
                AddItemsToKey(key, colStartIx, dataFieldItems, keys, 1d, CountValue);
            }
        }

		internal override void AggregateItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, List<bool> showTotals)
		{            
			if (value is KahanSum || IsNumeric(value))
			{
				AggregateKeys(key, colStartIx, dataFieldItems, keys, 1d, CountValue, showTotals);
			}
		}
	}

}
