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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime
{
    internal abstract class DateParsingFunction : ExcelFunction
    {
        protected DateTime ParseDate(IList<FunctionArgument> arguments, object dateObj, int argIndex, out ExcelErrorValue error)
        {
            error = null;
            DateTime date = DateTime.MinValue;
            if (dateObj is string)
            {
                date = DateTime.Parse(dateObj.ToString(), CultureInfo.CurrentCulture);
            }
            else
            {
                var d = ArgToDecimal(arguments, argIndex, out error);
                if (d >= 0)
                {
                    date = ConvertUtil.FromOADateExcel(d);
                }
            }
            return date;
        }

        protected DateTime ParseDate(IList<FunctionArgument> arguments, object dateObj, out ExcelErrorValue error)
        {
            return ParseDate(arguments, dateObj, 0, out error);
        }
    }
}
