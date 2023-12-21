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
            { eShowDataAs.Percent, new PivotShowAsPercent() },
            { eShowDataAs.PercentOfParentRow, new PivotShowAsPercentOfParentRowTotal()},
            { eShowDataAs.PercentOfParentColumn, new PivotShowAsPercentOfParentColumnTotal()},

            { eShowDataAs.RunningTotal, new PivotShowAsRunningTotal()},
        };
        internal static bool Calculate(ExcelPivotTable pivotTable, out List<Dictionary<int[], object>> calculatedItems, out List<Dictionary<int[], HashSet<int[]>>> keys)
        {
            var ci = pivotTable.CacheDefinition._cacheReference;
            calculatedItems = new List<Dictionary<int[], object>>();
            keys = new List<Dictionary<int[], HashSet<int[]>>>();
            var fieldIndex = pivotTable.RowColumnFieldIndicies;
            pivotTable.Filters.ReloadTable();
            foreach (var df in pivotTable.DataFields)
            {
                var dataFieldItems = PivotTableCalculation.GetNewCalculatedItems();
                var dataFieldKeys = PivotTableCalculation.GetNewKeys();
                calculatedItems.Add(dataFieldItems);
                var keyDict = new Dictionary<int[], HashSet<int[]>>(new ArrayComparer());
                keys.Add(keyDict);
                var recs = ci.Records;
                var captionFilters = pivotTable.Filters.Where(x => x.Type <= ePivotTableFilterType.ValueBetween).ToList();
				var pageFilterExists = pivotTable.PageFields.Count>0;
				var captionFilterExists = pivotTable.Filters.Count>0;
d
				for (var r = 0; r < recs.RecordCount; r++)
                {
                    var key = new int[fieldIndex.Count];
                    for (int i = 0; i < fieldIndex.Count; i++)
                    {
                        key[i] = (int)recs.CacheItems[fieldIndex[i]][r];
                    }
                    
                    if((pageFilterExists == false && PivotTableFilterMatcher.IsHiddenByPageField(pivotTable, recs, r) ||
					   (captionFilterExists == false && PivotTableFilterMatcher.IsHiddenByRowColumnFilter(pivotTable, captionFilters, recs, r))
                    {
                        _calculateFunctions[df.Function].AddItems(key, pivotTable.RowFields.Count, recs.CacheItems[df.Index][r], dataFieldItems, keyDict);
                    }
                }
                _calculateFunctions[df.Function].FilterValueFields(pivotTable, dataFieldItems);
				_calculateFunctions[df.Function].Calculate(recs.CacheItems[df.Index], dataFieldItems);
                if (df.ShowDataAs.Value != eShowDataAs.Normal)
                {
                    _calculateShowAs[df.ShowDataAs.Value].Calculate(df, fieldIndex, ref dataFieldItems);
                    calculatedItems[calculatedItems.Count - 1] = dataFieldItems;
                }
            }
            return true;
        }
    }
}
