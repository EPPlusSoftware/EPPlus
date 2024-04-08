/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/18/2024         EPPlus Software AB       EPPlus 7.1
*************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
	internal class PivotKeyUtil
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
			var end = (bf >= colFieldsStart ? fieldIndex.Count : colFieldsStart);
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