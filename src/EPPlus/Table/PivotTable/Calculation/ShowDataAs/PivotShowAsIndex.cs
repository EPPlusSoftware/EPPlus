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
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
	internal class PivotShowAsIndex : PivotShowAsBase
	{
		internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
		{
			var bf = fieldIndex.IndexOf(df.BaseField);
			var colFieldsStart = df.Field.PivotTable.RowFields.Count;
			var keyCol = fieldIndex.IndexOf(df.BaseField);
			var record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;
			var totKey = GetKey(fieldIndex.Count);
			var grandTotal = calculatedItems[totKey];
			if(grandTotal is not double grandNum)
			{
				grandNum = double.NaN;
			}
			var indexItems = new PivotCalculationStore();
			foreach (var key in calculatedItems.Index)
			{
				var cellValue = calculatedItems[key.Key];
				if(cellValue is double cellNumber)
				{
					var rowGrandKey = GetColumnTotalKey(key.Key, colFieldsStart);
					var rowGrand = calculatedItems[rowGrandKey];
					if (rowGrand is double rowNum)
					{
						var columnGrandKey = GetRowTotalKey(key.Key, colFieldsStart);
						var columnGrand = calculatedItems[columnGrandKey];
						if (columnGrand is double columnNum)
						{
							if(double.IsNaN(grandNum))
							{
								indexItems[key.Key] = grandTotal;
							}
							else
							{
								indexItems[key.Key] = cellNumber * grandNum / (rowNum * columnNum);
							}
						}
						else
						{
							indexItems[key.Key] = columnGrand;
						}
					}
					else
					{
						indexItems[key.Key] = rowGrand;
					}
				}
				else
				{
					indexItems[key.Key] = cellValue;
				}
			}
			calculatedItems = indexItems;
		}
	}
}
