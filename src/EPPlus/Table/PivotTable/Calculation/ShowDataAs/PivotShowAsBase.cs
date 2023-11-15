using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal abstract class PivotShowAsBase 
    {
        internal abstract void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref Dictionary<int[], object> calculatedItems);
        protected static int[] GetKey(int size, int iv = -1)
        {
            var key = new int[size];
            for (int i = 0; i < size; i++)
            {
                key[i] = iv;
            }
            return key;
        }

    }
}