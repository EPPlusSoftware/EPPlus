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
        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }
        public new bool Equals(object x, object y)
        {            
            if (x is string s1 && y is string s2)
            {
                return sc.Equals(s1, s2);
            }
            else
            {
                return object.Equals(GetValueToCompare(x), GetValueToCompare(y));
            }
        }

        public int GetHashCode(object obj)
        {
            return GetValueToCompare(obj).GetHashCode();
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        private static object GetValueToCompare(object obj)
        {
            var t = obj.GetType();
            var tc = Type.GetTypeCode(t);
            switch (tc)
            {
                case TypeCode.Empty:
                case TypeCode.Object:
                case TypeCode.Boolean:
                    return obj;
                case TypeCode.String:
                case TypeCode.Char:
                    return obj.ToString().ToLowerInvariant();
                case TypeCode.DateTime:
                    return ((DateTime)obj).ToOADate();
                default:
                    return Convert.ToDouble(obj);
            }
        }
    }
}