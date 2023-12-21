﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfParentColumnTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref Dictionary<int[], object> calculatedItems)
        {   
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var totalKey = GetKey(fieldIndex.Count);            
            var t = calculatedItems[totalKey];
            foreach(var key in calculatedItems.Keys.ToArray())
            {
                if (calculatedItems[key] is double d)
                {
                    var rowTotal = GetParentColumnTotal(key, colStartIx, calculatedItems, out ExcelErrorValue error);
                    if (double.IsNaN(rowTotal))
                    {
                        showAsCalculatedItems.Add(key,error);
                    }
                    else
                    {
                        showAsCalculatedItems.Add(key, d / rowTotal);
                    }                    
                }
            }
            calculatedItems = showAsCalculatedItems;
        }
        private static double GetParentColumnTotal(int[] key, int colStartIx, Dictionary<int[], object> calculatedItems, out ExcelErrorValue error)
        {
            var rowKey = (int[])key.Clone();
            for(int i=key.Length-1;i>=colStartIx;i--)
            {
                if(rowKey[i]!=-1)
                {
                    rowKey[i] = -1;
                    break;
                }
            }
            var v = calculatedItems[rowKey];
            if (v is ExcelErrorValue er)
            {
                error = er;
                return double.NaN;
            }
            error = null;
            return (double)v;
        }
    }
}