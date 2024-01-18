using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsRankAscending : PivotShowAsBase
    {
		internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems)
		{
			var bf = fieldIndex.IndexOf(df.BaseField);
			var colFieldsStart = df.Field.PivotTable.RowFields.Count;
			var keyCol = fieldIndex.IndexOf(df.BaseField);
			var record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;
			var maxBfKey = 0;
			var rankItems= new SortedList<int[], object>();
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
				else if (key.Key[keyCol]>=0)
				{
					if(IsSumAfter(key.Key, bf, fieldIndex, colFieldsStart, 1))
					{
						rankItems.Add(key.Key, calculatedItems[key.Key]);
					}
					else
					{
						calculatedItems[key.Key] = 1;
					}
				}
				else
				{
					calculatedItems[key.Key] = 1;
				}
			}

			while(rankItems.Count>0)
			{
				RankKeys(calculatedItems, rankItems, keyCol, maxBfKey);
			}
		}

		private void RankKeys(PivotCalculationStore calculatedItems, SortedList<int[], object> rankItems, int keyCol, int maxBfKey)
		{
			var key = rankItems.First();
			List<key>
			for(int i = 0; i < maxBfKey; i++)
			{

			}
		}
	}
}
