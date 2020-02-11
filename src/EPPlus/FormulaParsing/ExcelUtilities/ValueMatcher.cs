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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class ValueMatcher
    {
        public const int IncompatibleOperands = -2;

        public virtual int IsMatch(object o1, object o2)
        {
            if (o1 != null && o2 == null) return 1;
            if (o1 == null && o2 != null) return -1;
            if (o1 == null && o2 == null) return 0;
            //Handle ranges and defined names
            o1 = CheckGetRange(o1);
            o2 = CheckGetRange(o2);

            if (o1 is string && o2 is string)
            {
                return CompareStringToString(o1.ToString().ToLower(), o2.ToString().ToLower());
            }
            else if( o1.GetType() == typeof(string))
            {
                return CompareStringToObject(o1.ToString(), o2);
            }
            else if (o2.GetType() == typeof(string))
            {
                return CompareObjectToString(o1, o2.ToString());
            }
            else if(o1 is DateTime)
            {
                return ((DateTime)o1).ToOADate().CompareTo(Convert.ToDouble(o2));
            }
            return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
        }

        private static object CheckGetRange(object v)
        {
            if (v is ExcelDataProvider.IRangeInfo)
            {
                var r = ((ExcelDataProvider.IRangeInfo)v);
                if (r.GetNCells() > 1)
                {
                    v = ExcelErrorValue.Create(eErrorType.NA);
                }
                v = r.GetOffset(0, 0);
            }
            else if (v is ExcelDataProvider.INameInfo)
            {
                var n = ((ExcelDataProvider.INameInfo)v);
                v = CheckGetRange(n);
            }
            return v;
        }

        protected virtual int CompareStringToString(string s1, string s2)
        {
            return s1.CompareTo(s2);
        }

        protected virtual int CompareStringToObject(string o1, object o2)
        {
            double d1;
            if (double.TryParse(o1, out d1))
            {
                return d1.CompareTo(Convert.ToDouble(o2));
            }
            bool b1;
            if (bool.TryParse(o1, out b1))
            {
                return b1.CompareTo(Convert.ToBoolean(o2));
            }
            DateTime dt1;
            if (DateTime.TryParse(o1, out dt1))
            {
                return dt1.CompareTo(Convert.ToDateTime(o2));
            }
            return IncompatibleOperands;
        }

        protected virtual int CompareObjectToString(object o1, string o2)
        {
            double d2;
            if (double.TryParse(o2, out d2))
            {
                return Convert.ToDouble(o1).CompareTo(d2);
            }
            return IncompatibleOperands;
        }
    }
}
