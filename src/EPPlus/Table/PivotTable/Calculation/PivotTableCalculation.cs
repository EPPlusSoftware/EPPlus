﻿/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

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
            { DataFieldFunctions.StdDev,  new PivotFunctionStdDev() },
            { DataFieldFunctions.StdDevP,  null },
            { DataFieldFunctions.Var,  null },
            { DataFieldFunctions.VarP,  null }
        };
        internal static List<Dictionary<int[], object>> Calculate(ExcelPivotTable pivotTable)
        {
            var ci = pivotTable.CacheDefinition._cacheReference;
            var calculatedItems = new List<Dictionary<int[], object>>();
            var fieldIndex = pivotTable.RowColumnFieldIndicies;
            foreach (var df in pivotTable.DataFields)
            {
                var dataFieldItems = new Dictionary<int[], object>(new ArrayComparer());
                calculatedItems.Add(dataFieldItems);
                var recs = ci.Records;
                for (var r= 0; r < recs.RecordCount;r++)
                {
                    var key = new int[fieldIndex.Count];
                    for (int i=0;i < fieldIndex.Count;i++)
                    {
                        key[i] = (int)recs.CacheItems[fieldIndex[i]][r];
                    }

                    _calculateFunctions[df.Function].AddItems(key, recs.CacheItems[df.Index][r], dataFieldItems);
                }
                
                _calculateFunctions[df.Function].Calculate(recs.CacheItems[df.Index], dataFieldItems);
            }            
            return calculatedItems;
        }
    }
}