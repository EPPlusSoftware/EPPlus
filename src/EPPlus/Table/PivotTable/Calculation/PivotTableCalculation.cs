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
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using OfficeOpenXml.Table.PivotTable.Calculation.Filters;
using OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs;
using System;
using System.Collections.Generic;
using System.Linq;
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
            { eShowDataAs.Difference, new PivotShowAsDifference() },
            { eShowDataAs.PercentDifference, new PivotShowAsDifferencePercent() },
            { eShowDataAs.PercentOfParent, new PivotShowAsPercentOfParentTotal() },
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
			pivotTable.InitCalculation();
			int dfIx = 0;
			bool hasCalculatedField = false;
			foreach(var df in pivotTable.GetFieldsToCalculate())
            {
				var dataFieldItems = new PivotCalculationStore();
				calculatedItems.Add(dataFieldItems);
				var keyDict = PivotTableCalculation.GetNewKeys();
				keys.Add(keyDict);
				if (string.IsNullOrEmpty(df.Field.Cache.Formula))
				{
					CalculateField(pivotTable, calculatedItems[calculatedItems.Count - 1], keys, df.Field.Cache, df.Function);
                    SetRowColumnsItemsToHashSets(pivotTable);

                    if (df.ShowDataAs.Value != eShowDataAs.Normal)
					{
						_calculateShowAs[df.ShowDataAs.Value].Calculate(df, fieldIndex, keys[dfIx], ref dataFieldItems);
						calculatedItems[dfIx] = dataFieldItems;
					}
				}
				else
				{
					hasCalculatedField=true;
				}
				dfIx++;
			}

			CalculateRowColumnSubtotals(pivotTable, keys);

            //Handle Calculated fields after the pivot table fields has been calculated.
			if (hasCalculatedField)
			{
				CalculateSourceFields(pivotTable);
				var ptCalc = new PivotTableColumnCalculation(pivotTable);
				ptCalc.CalculateFormulaFields(fieldIndex);
			}

			return true;
        }

        private static void CalculateRowColumnSubtotals(ExcelPivotTable pivotTable, List<Dictionary<int[], HashSet<int[]>>> keys)
        {
			pivotTable.CalculatedFieldRowColumnSubTotals = new Dictionary<string, PivotCalculationStore>();
            foreach (var field in pivotTable.RowFields.Union(pivotTable.ColumnFields).Where(x=>x.SubTotalFunctions!=eSubTotalFunctions.None && x.SubTotalFunctions!=eSubTotalFunctions.Default))
			{
                var keyDict = PivotTableCalculation.GetNewKeys(); 
                keys.Add(keyDict);

				for(var dfIx=0;dfIx < pivotTable.DataFields.Count;dfIx++)
				{
					foreach (eSubTotalFunctions stf in Enum.GetValues(typeof(eSubTotalFunctions)))
					{
						if (stf == eSubTotalFunctions.None || stf == eSubTotalFunctions.Default) continue;
						if ((field.SubTotalFunctions & stf) != 0)
						{
							var store = new PivotCalculationStore();
							pivotTable.CalculatedFieldRowColumnSubTotals.Add($"{field.Index},{dfIx},{stf}", store);
							var df = pivotTable.DataFields[dfIx];
                            CalculateField(pivotTable, store, keys, df.Field.Cache, GetDataTypeFunction(stf));
						}
					}
				}
            }
        }

        private static DataFieldFunctions GetDataTypeFunction(eSubTotalFunctions stf)
        {
                switch (stf)
                {
                    case eSubTotalFunctions.Sum:
	                    return DataFieldFunctions.Sum;
                    case eSubTotalFunctions.Avg:
                        return DataFieldFunctions.Average;
                    case eSubTotalFunctions.Count:
                        return DataFieldFunctions.Count;
                    case eSubTotalFunctions.CountA:
                        return DataFieldFunctions.CountNums;
                    case eSubTotalFunctions.Product:
                        return DataFieldFunctions.Product;
                    case eSubTotalFunctions.Var:
                        return DataFieldFunctions.Var;
					case eSubTotalFunctions.VarP:
						return DataFieldFunctions.VarP;
					case eSubTotalFunctions.StdDev:
                        return DataFieldFunctions.StdDev;
					case eSubTotalFunctions.StdDevP:
						return DataFieldFunctions.StdDevP;
	                case eSubTotalFunctions.Min:
                        return DataFieldFunctions.Min;
                    case eSubTotalFunctions.Max:
                        return DataFieldFunctions.Max;
                    default:
                        return DataFieldFunctions.None;
                }
            }		
        private static void SetRowColumnsItemsToHashSets(ExcelPivotTable pivotTable)
        {
			if (pivotTable._colItems != null) return;
			var rowItems = new HashSet<int[]>(ArrayComparer.Instance);
            var colItems = new HashSet<int[]>(ArrayComparer.Instance);
			var rowLength = pivotTable.RowFields.Count;
			var colLength = pivotTable.ColumnFields.Count;
            foreach (var keyItem in pivotTable.CalculatedItems[0].Index)
			{
				if(rowLength>0)
				{
					var rowKey = new int[rowLength];
					Array.Copy(keyItem.Key,rowKey, rowLength);
					if(!rowItems.Contains(rowKey))
					{
						rowItems.Add(rowKey);
					}
                }
				if(colLength>0)
				{
                    var colKey = new int[colLength];
					Array.Copy(keyItem.Key, rowLength, colKey, 0, colLength);
					if (!colItems.Contains(colKey))
					{
						colItems.Add(colKey);
					}
                }
            }
            pivotTable._colItems = colItems;
            pivotTable._rowItems = rowItems;
        }

        private static void CalculateSourceFields(ExcelPivotTable pivotTable)
		{
			var keys = new List<Dictionary<int[], HashSet<int[]>>>();
			var calcFields = new Dictionary<string, PivotCalculationStore>(StringComparer.InvariantCultureIgnoreCase);
			foreach(var field in pivotTable.DataFields.Where(x=>string.IsNullOrEmpty(x.Field.Cache.Formula)==false).Select(x=>x.Field.Cache))
			{ 
				foreach(var token in field.FormulaTokens)
				{
					if(token.TokenType==TokenType.PivotField && calcFields.ContainsKey(token.Value)==false)
					{
						if(!GetSumCalcItems(pivotTable, token.Value, out PivotCalculationStore store))
						{
                            var keyDict = PivotTableCalculation.GetNewKeys();
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
			var pageFilterExists = pivotTable.PageFields.Select(x=>(x.MultipleItemSelectionAllowed && x.Items.HiddenItemIndex.Any()) || (x.MultipleItemSelectionAllowed==false && x.PageFieldSettings.SelectedItem>=0)).Count() > 0;
			var fieldIndex = pivotTable.RowColumnFieldIndicies;
			var slicerFields = pivotTable.Fields.Where(x=>x.Slicer!=null).ToList();
			var keyDict = keys[keys.Count-1];
			int index = cacheField.Index;
            for (var r = 0; r < recs.RecordCount; r++)
			{
				var key = new int[fieldIndex.Count];
				for (int i = 0; i < fieldIndex.Count; i++)
				{
					var field = pivotTable.Fields[fieldIndex[i]];
					if (field.Grouping == null)
					{
						key[i] = (int)recs.CacheItems[fieldIndex[i]][r];
					}
					else
					{
						int ix;
						if (field.Grouping.BaseIndex.HasValue && field.Grouping.BaseIndex != fieldIndex[i])
						{
							ix= field.Grouping.BaseIndex.Value;
						}
						else
						{
							ix = fieldIndex[i];
						}
						key[i] = field.GetGroupingKey((int)recs.CacheItems[ix][r]);						
					}
				 }

				if ((pageFilterExists == false || PivotTableFilterMatcher.IsHiddenByPageField(pivotTable, recs, r) == false) &&
					(captionFilterExists == false || PivotTableFilterMatcher.IsHiddenByRowColumnFilter(pivotTable, captionFilters, recs, r) == false) &&
					(slicerFields.Count == 0 || PivotTableFilterMatcher.IsHiddenBySlicer(pivotTable, recs, r, slicerFields)==false))
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
        internal static List<List<int[]>> GetAsCalculatedTable(ExcelPivotTable pivotTable)
        {
            var rowItems = pivotTable.GetTableRowKeys();
            var colItems = pivotTable.GetTableColumnKeys();
            var keyLength = pivotTable.RowFields.Count + pivotTable.ColumnFields.Count;
            var colStartIx = pivotTable.RowFields.Count;
            var table = new List<List<int[]>>();

            if (rowItems.Count == 0)
            {
                table.Add(colItems);
            }
            else if (colItems.Count == 0)
            {
				for (int r = 0; r < rowItems.Count; r++)
				{
					table.Add(new List<int[]> { rowItems[r] });
				}
            }
            else
            {
                for (int r = 0; r < rowItems.Count; r++)
                {
                    var l = new List<int[]>();
                    for (int c = 0; c < colItems.Count; c++)
                    {
                        var currentKey = new int[keyLength];
                        for (var i = 0; i < keyLength; i++)
                        {
                            if (i < colStartIx)
                            {
                                currentKey[i] = rowItems[r][i];
                            }
                            else
                            {
                                currentKey[i] = colItems[c][i - colStartIx];
                            }
                        }
                        l.Add(currentKey);
                    }
                    table.Add(l);
                }
            }
            return table;
        }

    }
}
