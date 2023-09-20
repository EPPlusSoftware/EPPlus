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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal static class DoubleArgParser
    {
        public static double Parse(object obj, out ExcelErrorValue error)
        {
            error = null;
            if (obj == null)
            {
                error = ExcelErrorValue.Create(eErrorType.Value);
                return double.NaN;
            }
            if (obj is IRangeInfo)
            {
                var r = ((IRangeInfo)obj).FirstOrDefault();
                return r == null ? 0 : r.ValueDouble;
            }
            if (obj is double dRes) return dRes;
            if (obj.IsNumeric()) return ConvertUtil.GetValueDouble(obj);
            var str = obj.ToString();
            try
            {
                double d;
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                    return d;
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;

                return DateTime.Parse(str, CultureInfo.CurrentCulture, DateTimeStyles.None).ToOADate();
            }
            catch// (Exception e)
            {
                error = ExcelErrorValue.Create(eErrorType.Value);
                return double.NaN;
            }
        }
    }
}
