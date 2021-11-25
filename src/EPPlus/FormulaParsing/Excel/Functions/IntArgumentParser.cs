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
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class IntArgumentParser : ArgumentParser
    {
        public override object Parse(object obj)
        {
            return Parse(obj, RoundingMethod.Convert);
        }

        public override object Parse(object obj, RoundingMethod roundingMethod)
        {
            Require.That(obj).Named("argument").IsNotNull();
            int result;
            if (obj is ExcelDataProvider.IRangeInfo)
            {
                var r = ((ExcelDataProvider.IRangeInfo)obj).FirstOrDefault();
                return r == null ? 0 : ConvertToInt(r.ValueDouble, roundingMethod);
            }
            var objType = obj.GetType();
            if (objType == typeof(int))
            {
                return (int)obj;
            }
            if (objType == typeof(double) || objType == typeof(decimal) || objType == typeof(bool))
            {
                return ConvertToInt(obj, roundingMethod);
            }
            if (!int.TryParse(obj.ToString(), out result))
            {
                throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
            }
            return result;
        }

        private int ConvertToInt(object obj, RoundingMethod roundingMethod)
        {
            var objType = obj.GetType();
            if (roundingMethod == RoundingMethod.Convert)
            {
                return Convert.ToInt32(obj);
            }
            else if (objType == typeof(double))
            {
                return Convert.ToInt32(System.Math.Floor((double)obj));
            }
            else
            {
                return Convert.ToInt32(System.Math.Floor((decimal)obj));
            }
        }
    }
}
