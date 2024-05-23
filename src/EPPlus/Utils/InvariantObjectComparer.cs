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
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Utils
{
    internal class InvariantObjectComparer : IEqualityComparer<object>
    {
		internal static InvariantObjectComparer Instance=new InvariantObjectComparer();

        static StringComparer sc = StringComparer.OrdinalIgnoreCase;

        public new bool Equals(object x, object y)
        {            
            if (x is string s1 && y is string s2)
            {
                return sc.Equals(s1, s2);
            }
            else
            {
                return object.Equals(x, y);
            }
        }

        public int GetHashCode(object obj)
        {
            if (obj is string s)
            {
                return s.ToLowerInvariant().GetHashCode();
            }
            
            return obj.GetHashCode();
        }
    }
}