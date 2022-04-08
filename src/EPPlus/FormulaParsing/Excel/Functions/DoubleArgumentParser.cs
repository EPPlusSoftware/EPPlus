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
using System.Globalization;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using util=OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class DoubleArgumentParser : ArgumentParser
    {
        public override object Parse(object obj)
        {
            Require.That(obj).Named("argument").IsNotNull();
            if (obj is IRangeInfo)
            {
                var r=((IRangeInfo)obj).FirstOrDefault();
                return r == null ? 0 : r.ValueDouble;
            }
            if (obj is double) return obj;
            if (obj.IsNumeric()) return util.ConvertUtil.GetValueDouble(obj);
            var str = obj != null ? obj.ToString() : string.Empty;
            try
            {
                double d;
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;

                return System.DateTime.Parse(str, CultureInfo.CurrentCulture, DateTimeStyles.None).ToOADate();
            }
            catch// (Exception e)
            {
                throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
            }
        }

        public override object Parse(object obj, RoundingMethod roundingMethod)
        {
            return Parse(obj);
        }
    }
}
