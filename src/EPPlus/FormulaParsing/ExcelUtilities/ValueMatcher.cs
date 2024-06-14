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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Compares and matches values
    /// </summary>
    public class ValueMatcher
    {
        /// <summary>
        /// Value to represent incompatible operands
        /// </summary>
        public const int IncompatibleOperands = -2;

        /// <summary>
        /// Compares objects of different types using appropriate CompareTo methods
        /// </summary>
        /// <param name="searchedValue">original value</param>
        /// <param name="candidate">potential match</param>
        /// <returns></returns>
        public virtual int IsMatch(object searchedValue, object candidate)
        {
            if (searchedValue != null && candidate == null) return -1;
            if (searchedValue == null && candidate != null) return 1;
            if (searchedValue == null && candidate == null) return 0;
            //Handle ranges and defined names
            searchedValue = CheckGetRange(searchedValue);
            candidate = CheckGetRange(candidate);

            if ((searchedValue is string && candidate is string) ||
                (candidate.GetType() == typeof(ExcelErrorValue)))
            {
                return CompareStringToString(searchedValue.ToString().ToLower(), candidate.ToString().ToLower());
            }
            else if(searchedValue.GetType() == typeof(string))
            {
                return CompareStringToObject(searchedValue.ToString(), candidate);
            }
            else if (candidate.GetType() == typeof(string))
            {
                return CompareObjectToString(searchedValue, candidate.ToString());
            }
            else if(candidate is DateTime && searchedValue is DateTime)
            {
                return ((DateTime)candidate).CompareTo(((DateTime)searchedValue));
            }
            else if(candidate is DateTime)
            {
                return ((DateTime)candidate).ToOADate().CompareTo(Convert.ToDouble(searchedValue));
            }
            else if(searchedValue is DateTime)
            {
                return Convert.ToDouble(candidate).CompareTo(((DateTime)searchedValue).ToOADate());
            }
            else if (candidate is IConvertible && searchedValue is IConvertible)
            {
                return Convert.ToDouble(candidate).CompareTo(Convert.ToDouble(searchedValue));
            }
            else
            {
                return IncompatibleOperands;
            }
        }

        private static object CheckGetRange(object v)
        {
            if (v is IRangeInfo)
            {
                var r = ((IRangeInfo)v);
                if (r.GetNCells() > 1)
                {
                    v = ExcelErrorValue.Create(eErrorType.NA);
                }
                v = r.GetOffset(0, 0);
            }
            else if (v is INameInfo)
            {
                var n = ((INameInfo)v);
                v = CheckGetRange(n);
            }
            return v;
        }
        /// <summary>
        /// Compares strings
        /// </summary>
        /// <param name="searchedValue"></param>
        /// <param name="candidate"></param>
        /// <returns></returns>
        protected virtual int CompareStringToString(string searchedValue, string candidate)
        {
            return candidate.CompareTo(searchedValue);
        }
        /// <summary>
        /// Compares string to object candidate
        /// </summary>
        /// <param name="searchedValue"></param>
        /// <param name="candidate"></param>
        /// <returns></returns>
        protected virtual int CompareStringToObject(string searchedValue, object candidate)
        {
            if (double.TryParse(searchedValue, out double dsv))
            {
                return ConvertUtil.GetValueDouble(candidate).CompareTo(dsv);
            }
            if (bool.TryParse(searchedValue, out bool bsv))
            {
                if(candidate is bool cb)
                {
                    return cb.CompareTo(bsv);
                }                
            }
            if (DateTime.TryParse(searchedValue, out DateTime dtsv))
            {
                DateTime? date = ConvertUtil.GetValueDate(candidate);
                if (date.HasValue == false)
                    return -1;
                return date.Value.CompareTo(dtsv);
            }
            return IncompatibleOperands;
        }
        /// <summary>
        /// Compares object to string candidate.
        /// </summary>
        /// <param name="searchedValue"></param>
        /// <param name="candidate"></param>
        /// <returns></returns>
        protected virtual int CompareObjectToString(object searchedValue, string candidate)
        {
            if (double.TryParse(candidate, out double d2))
            {
                return d2.CompareTo(Convert.ToDouble(searchedValue));
            }
            return IncompatibleOperands;
        }
    }
}
