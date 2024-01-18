using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal abstract class PivotShowAsBase 
    {
        internal abstract void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems);
        protected static int[] GetKey(int size, int iv = PivotCalculationStore.SumLevelValue)
        {
            var key = new int[size];
            for (int i = 0; i < size; i++)
            {
                key[i] = iv;
            }
            return key;
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