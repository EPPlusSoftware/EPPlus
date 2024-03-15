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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
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
			calculatedItems = new List<PivotCalculationStore>();
			keys = new List<Dictionary<int[], HashSet<int[]>>>();
            var fieldIndex = pivotTable.RowColumnFieldIndicies;
            pivotTable.Filters.ReloadTable();
			pivotTable.Sort();
			int dfIx = 0;
			bool hasCf = false;
			foreach (var df in pivotTable.GetFieldsToCalculate())
            {
				dfIx++;
				var dataFieldItems = new PivotCalculationStore();
				calculatedItems.Add(new PivotCalculationStore());
				var keyDict = new Dictionary<int[], HashSet<int[]>>(new ArrayComparer());
				keys.Add(keyDict);
				if (string.IsNullOrEmpty(df.Field.CacheField.Formula))
				{
					CalculateField(pivotTable, calculatedItems[calculatedItems.Count-1], keys, df.Field.CacheField, df.Function);

					if (df.ShowDataAs.Value != eShowDataAs.Normal)
					{
						_calculateShowAs[df.ShowDataAs.Value].Calculate(df, fieldIndex, ref dataFieldItems);
						calculatedItems[dfIx] = dataFieldItems;
					}
				}
				else
				{
					hasCf=true;
				}
            }
			
			if(hasCf)
			{
				CalculateSourceFields(pivotTable);
				var ptCalc = new PivotTableColumnCalculation(pivotTable);
				ptCalc.CalculateFormulaFields(fieldIndex);
			}
			
			return true;
        }

		private static void CalculateSourceFields(ExcelPivotTable pivotTable)
		{
			var keys = new List<Dictionary<int[], HashSet<int[]>>>();
			var calcFields = new Dictionary<string, PivotCalculationStore>(StringComparer.InvariantCultureIgnoreCase);
			foreach(var field in pivotTable.DataFields.Where(x=>string.IsNullOrEmpty(x.Field.CacheField.Formula)==false).Select(x=>x.Field.CacheField))
			{ 
				foreach(var token in field.FormulaTokens)
				{
					if(token.TokenType==TokenType.PivotField && calcFields.ContainsKey(token.Value)==false)
					{
						if(!GetSumCalcItems(pivotTable, token.Value, out PivotCalculationStore store))
						{
							var keyDict = new Dictionary<int[], HashSet<int[]>>(new ArrayComparer());
							keys.Add(keyDict);
							store = new PivotCalculationStore();
							CalculateField(pivotTable, store, keys, pivotTable.Fields[token.Value].Cache, DataFieldFunctions.Sum);
						}
						calcFields.Add(token.Value, store);
					}
				}
			}
			pivotTable.CalculatedFieldReferencedItems = calcFields;
		}

		private static bool GetSumCalcItems(ExcelPivotTable pivotTable, string fieldName, out PivotCalculationStore store)
		{
			foreach(var ds in pivotTable.DataFields)
			{
				if(ds.Field!=null && ds.Field.Name.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase) && ds.Function==DataFieldFunctions.Sum && ds.ShowDataAsInternal == eShowDataAs.Normal)
				{
					store = pivotTable.CalculatedItems[ds.Index];
					return true;
				}
			}
			store = null;
			return false;
		}

		private static void CalculateField(ExcelPivotTable pivotTable, PivotCalculationStore dataFieldItems, List<Dictionary<int[], HashSet<int[]>>> keys,  ExcelPivotTableCacheField cacheField, DataFieldFunctions function)
		{
			var ci = pivotTable.CacheDefinition._cacheReference;
			var recs = ci.Records;
			var captionFilters = pivotTable.Filters.Where(x => x.Type < ePivotTableFilterType.ValueBetween).ToList();
			var captionFilterExists = captionFilters.Count > 0;
			var pageFilterExists = pivotTable.PageFields.Count > 0;
			var fieldIndex = pivotTable.RowColumnFieldIndicies;
			var keyDict = keys[keys.Count-1];
			int index = cacheField.Index;
			
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
					var v = cacheField.IsRowColumnOrPage ? cacheField.SharedItems[(int)recs.CacheItems[index][r]] : recs.CacheItems[index][r];
					_calculateFunctions[function].AddItems(key, pivotTable.RowFields.Count, v, dataFieldItems, keyDict);
				}
			}

			_calculateFunctions[function].FilterValueFields(pivotTable, dataFieldItems, keys[keys.Count - 1], fieldIndex);
			_calculateFunctions[function].Aggregate(pivotTable, dataFieldItems, keys[keys.Count - 1]);
			_calculateFunctions[function].Calculate(recs.CacheItems[index], dataFieldItems);
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
