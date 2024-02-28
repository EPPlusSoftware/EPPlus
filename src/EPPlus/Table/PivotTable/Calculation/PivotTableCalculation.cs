/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2024         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using EPPlusTest.Table.PivotTable;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using OfficeOpenXml.Table.PivotTable.Calculation.Filters;
using OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using OfficeOpenXml.Table.PivotTable.Calculation;
namespace OfficeOpenXml.Table.PivotTable
{
    internal partial class PivotTableCalculation
    {
        static Dictionary<DataFieldFunctions, PivotFunction> _calculateFunctions = new Dictionary<DataFieldFunctions, PivotFunction>
        {
            { DataFieldFunctions.None, new PivotFunctionSum() },
            { DataFieldFunctions.Sum, new PivotFunctionSum() },
            { DataFieldFunctions.Count, new PivotFunctionCount() },
            { DataFieldFunctions.CountNums, new PivotFunctionCountNums() },
            { DataFieldFunctions.Product, new PivotFunctionProduct() },
            { DataFieldFunctions.Min, new PivotFunctionMin() },
            { DataFieldFunctions.Max, new PivotFunctionMax() },
            { DataFieldFunctions.Average, new PivotFunctionAverage() },
            { DataFieldFunctions.StdDev,  new PivotFunctionStdDev() },
            { DataFieldFunctions.StdDevP,  new PivotFunctionStdDevP() },
            { DataFieldFunctions.Var,  new PivotFunctionVar() },
            { DataFieldFunctions.VarP,  new PivotFunctionVarP() }
        };

        static Dictionary<eShowDataAs, PivotShowAsBase> _calculateShowAs = new Dictionary<eShowDataAs, PivotShowAsBase>
        {
            { eShowDataAs.PercentOfTotal, new PivotShowAsPercentOfGrandTotal() },
            { eShowDataAs.PercentOfColumn, new PivotShowAsPercentOfColumnTotal() },
            { eShowDataAs.PercentOfRow, new PivotShowAsPercentOfRowTotal() },
            { eShowDataAs.Percent, new PivotShowAsPercentOf() },
            { eShowDataAs.PercentOfParentRow, new PivotShowAsPercentOfParentRowTotal()},
            { eShowDataAs.PercentOfParentColumn, new PivotShowAsPercentOfParentColumnTotal()},
            { eShowDataAs.RunningTotal, new PivotShowAsRunningTotal()},
			{ eShowDataAs.PercentOfRunningTotal, new PivotShowAsPercentOfRunningTotal()},
			{ eShowDataAs.RankAscending, new PivotShowAsRankAscending()},
            { eShowDataAs.RankDescending, new PivotShowAsRankDescending()},
			{ eShowDataAs.Index, new PivotShowAsIndex()},
		};
        internal static bool Calculate(ExcelPivotTable pivotTable, out List<PivotCalculationStore> calculatedItems, out List<Dictionary<int[], HashSet<int[]>>> keys)
        {
            var ci = pivotTable.CacheDefinition._cacheReference;
			calculatedItems = new List<PivotCalculationStore>();
			keys = new List<Dictionary<int[], HashSet<int[]>>>();
            var fieldIndex = pivotTable.RowColumnFieldIndicies;
            pivotTable.Filters.ReloadTable();
            foreach (var df in pivotTable.DataFields)
            {
                var dataFieldItems = new PivotCalculationStore();
                calculatedItems.Add(dataFieldItems);
                var keyDict = new Dictionary<int[], HashSet<int[]>>(new ArrayComparer());
                keys.Add(keyDict);
				var recs = ci.Records;
                var captionFilters = pivotTable.Filters.Where(x => x.Type < ePivotTableFilterType.ValueBetween).ToList();
				var captionFilterExists = captionFilters.Count > 0;
				var pageFilterExists = pivotTable.PageFields.Count > 0;
				var cacheField = df.Field.Cache;

				for (var r = 0; r < recs.RecordCount; r++)
                {
                    var key = new int[fieldIndex.Count];
                    for (int i = 0; i < fieldIndex.Count; i++)
                    {
						if (pivotTable.Fields[fieldIndex[i]].Grouping == null)
						{
							key[i] = (int)recs.CacheItems[fieldIndex[i]][r];
						}
						else
						{
							key[i] = pivotTable.Fields[fieldIndex[i]].GetGroupingKey((int)recs.CacheItems[fieldIndex[i]][r]);
						}
                    }

                    if ((pageFilterExists == false || PivotTableFilterMatcher.IsHiddenByPageField(pivotTable, recs, r) == false) &&
                        (captionFilterExists == false || PivotTableFilterMatcher.IsHiddenByRowColumnFilter(pivotTable, captionFilters, recs, r) == false))
                    {
						var v = cacheField.IsRowColumnOrPage ? cacheField.SharedItems[(int)recs.CacheItems[df.Index][r]] : recs.CacheItems[df.Index][r];
						_calculateFunctions[df.Function].AddItems(key, pivotTable.RowFields.Count, v, dataFieldItems, keyDict);
                    }
                }

				_calculateFunctions[df.Function].FilterValueFields(pivotTable, dataFieldItems, keys[calculatedItems.Count - 1], fieldIndex);
				_calculateFunctions[df.Function].Aggregate(pivotTable, dataFieldItems, keys[calculatedItems.Count-1]);
				_calculateFunctions[df.Function].Calculate(recs.CacheItems[df.Index], dataFieldItems);

                if (df.ShowDataAs.Value != eShowDataAs.Normal)
                {
                    _calculateShowAs[df.ShowDataAs.Value].Calculate(df, fieldIndex, ref dataFieldItems);
                    calculatedItems[calculatedItems.Count - 1] = dataFieldItems;
                }
            }
            return true;
        }
		internal static int[] GetKeyWithParentLevel(int[] key, int[] childKey, int rf)
		{
			if (IsKeyGrandTotal(key, 0, rf) == false)
			{
				for (var i = 0; i < rf; i++)
				{
					if (key[i] == PivotCalculationStore.SumLevelValue)
					{
						key[i] = childKey[i];
					}
					else
					{
						break;
					}
				}
			}
			if (IsKeyGrandTotal(key, rf, childKey.Length) == false)
			{
				for (var i = rf; i <= key.Length - 1; i++)
				{
					if (key[i] == PivotCalculationStore.SumLevelValue)
					{
						key[i] = childKey[i];
					}
					else
					{
						break;
					}
				}
			}
			return key;
		}
		internal static bool IsKeyGrandTotal(int[] key, int startIx, int endIx)
		{
			for (int i = startIx; i < endIx; i++)
			{
				if (key[i] != PivotCalculationStore.SumLevelValue)
				{
					return false;
				}
			}

			return true;
		}
		internal static bool IsReferencingUngroupableKey(int[] key, int rf)
		{
			for (var i = 1; i < rf; i++)
			{
				if (key[i - 1] == PivotCalculationStore.SumLevelValue && key[i] != PivotCalculationStore.SumLevelValue)
				{
					return true;
				}
			}

			for (var i = rf + 1; i <= key.Length - 1; i++)
			{
				if (key[i - 1] == PivotCalculationStore.SumLevelValue && key[i] != PivotCalculationStore.SumLevelValue)
				{
					return true;
				}
			}
			return false;
		}
	}
}
