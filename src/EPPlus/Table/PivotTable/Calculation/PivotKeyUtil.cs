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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Calculation
{
	internal abstract class PivotKeyUtil
	{
		internal static int[] GetKey(int size, int iv = PivotCalculationStore.SumLevelValue)
		{
			var key = new int[size];
			for (int i = 0; i < size; i++)
			{
				key[i] = iv;
			}
			return key;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="key"></param>
		/// <param name="colFieldsStart">Where row fields end and colfields start in the key</param>
		/// <returns></returns>
		internal static int[] GetColumnTotalKey(int[] key, int colFieldsStart)
		{
			var newKey = (int[])key.Clone();
			for (int i = 0; i < colFieldsStart; i++)
			{
				newKey[i] = PivotCalculationStore.SumLevelValue;
			}
			return newKey;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="key"></param>
		/// <param name="colFieldsStart">Where row fields end and colfields start in the key</param>
		/// <returns></returns>
		internal static int[] GetRowTotalKey(int[] key, int colFieldsStart)
		{
			var newKey = (int[])key.Clone();
			for (int i = colFieldsStart; i < newKey.Length; i++)
			{
				newKey[i] = PivotCalculationStore.SumLevelValue;
			}
			return newKey;
		}
		internal static int[] GetParentKey(int[] key, int keyCol)
		{
			var newKey = (int[])key.Clone();
			newKey[keyCol] = PivotCalculationStore.SumLevelValue;
			return newKey;
		}
		internal static int[] GetNextKey(int[] key, int keyCol)
		{
			var newKey = (int[])key.Clone();
			newKey[keyCol]++;
			return newKey;
		}
		internal static int[] GetPrevKey(int[] key, int keyCol)
		{
			var newKey = (int[])key.Clone();
			newKey[keyCol]--;
			return newKey;
		}

		internal static bool IsSumBefore(int[] key, int bf, List<int> fieldIndex, int colFieldsStart)
		{
			var start = (bf >= colFieldsStart ? colFieldsStart : 0);
			for (int i = start; i <= bf; i++)
			{
				if (key[i] == PivotCalculationStore.SumLevelValue)
				{
					return true;
				}
			}
			return false;
		}
		internal static bool IsSumAfter(int[] key, int bf, List<int> fieldIndex, int colFieldsStart, int addStart = 0)
		{
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
        internal static bool IsRowGrandTotal(int[] key, int colFieldStart)
		{
			for(int i = 0; i < colFieldStart; i++)
			{
				if (key[i]!=PivotCalculationStore.SumLevelValue)
				{
					return false;
				}
			}
			return true;
		}
        internal static bool IsColumnGrandTotal(int[] key, int colFieldStart)
        {
            for (int i = colFieldStart; i < key.Length; i++)
            {
                if (key[i] != PivotCalculationStore.SumLevelValue)
                {
                    return false;
                }
            }
            return true;
        }

        internal static int[] GetKeyPart(int[] key, int fromIndex, int toIndex)
        {
			var newKey = new int[key.Length];

            for (int i=0;i<key.Length;i++)
			{
				if(i>=fromIndex && i<= toIndex)
				{
					newKey[i] = key[i];	
				}
				else
				{
                    newKey[i] = PivotCalculationStore.SumLevelValue;
                }
			}
			return newKey;
        }

        protected static int[] GetNextKeyFromKeys(HashSet<int[]> keysParent, int keyCol, int index)
        {
            int[] ret = null;
			foreach(var k in keysParent)
			{
				if (k[keyCol] > index)
				{
					if (ret == null || ret[keyCol] > k[keyCol])
					{
						ret = k;
					}
				}
			}
			return ret;
        }
        protected static int[] GetPreviousKeyFromKeys(HashSet<int[]> keysParent, int keyCol, int index)
        {
            int[] ret = null;
            foreach (var k in keysParent)
            {
                if (k[keyCol] < index)
                {
                    if (ret == null || ret[keyCol] < k[keyCol])
                    {
                        ret = k;
                    }
                }
            }
            return ret;
        }
        internal static bool ExistsValueInTable(int[] key, int colFieldStart, PivotCalculationStore calculatedItems)
        {
            var rowKey = PivotKeyUtil.GetRowTotalKey(key, colFieldStart);
            var colKey = PivotKeyUtil.GetColumnTotalKey(key, colFieldStart);
            return calculatedItems.ContainsKey(rowKey) && calculatedItems.ContainsKey(colKey);
        }
        protected static int[] GetPrevKeyFromCalculatedTable(List<List<int[]>> calcTable, int r, int c, int keyCol, bool isRowField)
        {
            if (isRowField)
            {
                if (r == 0)
                {
                    return null;
                }
                else
                {
                    var pr = r;
                    
                    while (--pr >= 0 && HasSameParent(calcTable[r][c], calcTable[pr][c], keyCol) == false);
                    
                    if(pr>=0)
                    {
                        return calcTable[pr][c];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            else
            {
                if (c == 0)
                {
                    return null;
                }
                else
                {
                    var pc = c;

                    while (--pc >= 0 && HasSameParent(calcTable[r][c], calcTable[r][pc], keyCol) == false) ;
                    if (pc >= 0)
                    {
                        return calcTable[r][pc];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }
        protected static int[] GetNextKeyFromCalculatedTable(List<List<int[]>> calcTable, int r, int c, int keyCol, bool isRowField)
        {
            if (isRowField)
            {
                if (r == calcTable.Count - 1)
                {
                    return null;
                }
                else
                {
                    var pr = r;

                    while (++pr < calcTable.Count && HasSameParent(calcTable[r][c], calcTable[pr][c], keyCol) == false) ;

                    if (pr < calcTable.Count)
                    {
                        return calcTable[pr][c];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            else
            {
                if (c == calcTable[r].Count - 1)
                {
                    return null;
                }
                else
                {
                    var pc = c;

                    while (++pc < calcTable.Count && HasSameParent(calcTable[r][c], calcTable[r][pc], keyCol) == false);
                    if (pc >= 0)
                    {
                        return calcTable[r][pc];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        protected static bool HasSameParent(int[] key1, int[] key2, int keyCol)
        {
            for (int i = 0; i < key1.Length; i++)
            {
                if (i != keyCol && key1[i] != key2[i])
                {
                    return false;
                }
            }
            return true;
        }

        protected static bool IsSameLevelAs(int[] key, bool isRowField, int baseLevel, int keyCol, ExcelPivotTableDataField df)
        {
            if (isRowField)
            {
                for (int i = baseLevel + 1; i < df.Field.PivotTable.RowFields.Count; i++)
                {
                    if (key[i] != PivotCalculationStore.SumLevelValue)
                    {
                        return false;
                    }
                }
                return true;
            }
            else
            {
                for (int i = baseLevel + 1; i < df.Field.PivotTable.ColumnFields.Count; i++)
                {
                    if (key[i] != PivotCalculationStore.SumLevelValue)
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}