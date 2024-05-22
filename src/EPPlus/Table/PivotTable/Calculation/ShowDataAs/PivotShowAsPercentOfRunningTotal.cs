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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfRunningTotal : PivotShowAsRunningTotal
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {
            var calcPercent = PivotTableCalculation.GetNewCalculatedItems();
            if (fieldIndex.Count == 0)
            {
                calculatedItems.Values[0]=ErrorValues.NAError;
                return;
            }
            var colFieldsStart = df.Field.PivotTable.RowFields.Count;
            var keyCol = fieldIndex.IndexOf(df.BaseField);
            var record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;

            CalculateRunningTotal(df, fieldIndex, keys, ref calculatedItems, true);

            foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
            {
                if (IsSumBefore(key.Key, keyCol, fieldIndex, colFieldsStart))
                {
                    calcPercent[key.Key] = 0D;
                }
                else if(ShouldCalculateKey(key.Key, keyCol, fieldIndex, colFieldsStart))
                {
                    var rtValue = calculatedItems[key.Key];
                    if (rtValue is double value)
                    {
                        var sumKey = (int[])key.Key.Clone();
                        sumKey[keyCol] = PivotCalculationStore.SumLevelValue;
                        var rtSum = calculatedItems[sumKey];
                        if(rtSum is double sumValue)
                        {
                            if(sumValue==0D && value != 0D)
                            {
                                calcPercent[key.Key] = ErrorValues.Div0Error;
                            }
                            else
                            {
                                calcPercent[key.Key] = value / sumValue;
                            }
                        }
                        else
                        {
                            calcPercent[key.Key] = rtSum;
                        }
                    }
                    else
                    {   
                        calcPercent[key.Key] = rtValue;
                    }
                }
                else
                {
                    var rtValue = calculatedItems[key.Key];
                    if (rtValue == null)
                    {
                        calcPercent[key.Key] = 0D;
                    }
                    else if (rtValue is double d)
                    {
                        if (d == 0D)
                        {
                            calcPercent[key.Key] = 0D;
                        }
                        else
                        {
                            calcPercent[key.Key] = 1D;
                        }
                    }
                    else
                    {

                        calcPercent[key.Key] = rtValue;
                    }
                }
            }
            calculatedItems = calcPercent;
        }
        internal static bool ShouldCalculateKey(int[] key, int bf, List<int> fieldIndex, int colFieldsStart, int addStart = 0)
        {
            if(key[bf] == PivotCalculationStore.SumLevelValue) return false;
            if(bf == colFieldsStart - 1 || bf == fieldIndex.Count - 1) return true;
            var start = bf + addStart;
            var end = (bf >= colFieldsStart ? fieldIndex.Count : colFieldsStart);
            if (start == end) return false;
            for (int i = bf + addStart; i < end; i++)
            {
                if (key[i] == PivotCalculationStore.SumLevelValue)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
