/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.Table.PivotTable.Calculation;
using System.Collections.Generic;
using System.Diagnostics;

namespace OfficeOpenXml.Table.PivotTable
{
    internal partial class PivotTableCalculation
    {
		internal static PivotCalculationStore GetNewCalculatedItems()
		{
			return new PivotCalculationStore();
		}
		internal static Dictionary<int[], HashSet<int[]>> GetNewKeys()
        {
            return new Dictionary<int[], HashSet<int[]>>(new ArrayComparer());
        }
    }
    internal class ArrayComparer : IEqualityComparer<int[]>, IComparer<int[]>
    {
        internal static readonly ArrayComparer Instance = new ArrayComparer();
        public static bool IsEqual(int[] x, int[] y)
        {
            if (x.Length != y.Length) return false;
            for (int i = 0; i < x.Length; i++)
            {
                if (x[i] != y[i]) return false;
            }
            return true;
        }

        public int Compare(int[] x, int[] y)
        {
            for(int i=0;i<x.Length;i++)
            {
                if (x[i] != y[i])
                { 
                    return x[i].CompareTo(y[i]); 
                } 
            }
            return 0;
        }

        public bool Equals(int[] x, int[] y)
        {
            return IsEqual(x, y);
        }

        public int GetHashCode(int[] obj)
        {
            int hash = 49;
            for (int i = 1; i < obj.Length; i++)
            {
                unchecked
                {
                    hash *= 23 * obj[i].GetHashCode();
                }
            }
            return hash;
        }
    }

}