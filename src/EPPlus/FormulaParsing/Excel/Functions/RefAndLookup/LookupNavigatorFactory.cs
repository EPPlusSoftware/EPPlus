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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal static class LookupNavigatorFactory
    {
        public static LookupNavigator Create(LookupDirection direction, LookupArguments args, ParsingContext parsingContext)
        {
            if (args.ArgumentDataType == LookupArguments.LookupArgumentDataType.ExcelRange)
            {
                return new ExcelLookupNavigator(direction, args, parsingContext);
            }
            else if (args.ArgumentDataType == LookupArguments.LookupArgumentDataType.DataArray)
            {
                return new ArrayLookupNavigator(direction, args, parsingContext);
            }
            throw new NotSupportedException("Invalid argument datatype");
        }
    }
}
