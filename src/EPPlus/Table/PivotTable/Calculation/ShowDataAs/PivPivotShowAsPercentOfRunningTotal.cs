using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
	internal class PivotShowAsPercentOfRunningTotal : PivotShowAsRunningTotal
	{
		internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems)
		{
			var bf = fieldIndex.IndexOf(df.BaseField);
			var colFieldsStart = df.Field.PivotTable.RowFields.Count;
			var keyCol = fieldIndex.IndexOf(df.BaseField);
			var record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;
			var maxBfKey = 0;

			CalculateRunningTotal(df, fieldIndex, ref calculatedItems, true);

			if (record.CacheItems[df.BaseField].Count(x => x is int) > 0)
			{
				maxBfKey = (int)record.CacheItems[df.BaseField].Where(x => x is int).Max();
			}

			foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
			{
				if (IsSumBefore(key.Key, bf, fieldIndex, colFieldsStart))
				{
					calculatedItems[key.Key] = null;
				}
				else
				{
					if (IsSumAfter(key.Key, bf, fieldIndex, colFieldsStart))
					{
						var parentKey = GetParentKey(key.Key, keyCol);
						var parentValue = calculatedItems[parentKey];
						if(parentValue is double pv)
						{
							calculatedItems[key.Key] = (double)calculatedItems[key.Key] / pv;
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
