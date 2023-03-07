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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    internal abstract class DateParsingFunction : ExcelFunction
    {
        protected System.DateTime ParseDate(IEnumerable<FunctionArgument> arguments, object dateObj, int argIndex)
        {
            System.DateTime date = System.DateTime.MinValue;
            if (dateObj is string)
            {
                date = System.DateTime.Parse(dateObj.ToString(), CultureInfo.InvariantCulture);
            }
            else
            {
                var d = ArgToDecimal(arguments, argIndex);
                if (d >= 0)
                {
                    date = ConvertUtil.FromOADateExcel(d);
                }
            }
            return date;
        }

        protected System.DateTime ParseDate(IEnumerable<FunctionArgument> arguments, object dateObj)
        {
            return ParseDate(arguments, dateObj, 0);
        }
    }
}
