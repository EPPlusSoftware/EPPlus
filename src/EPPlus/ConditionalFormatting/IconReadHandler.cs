﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Rules;
using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal static class IconReadHandler
    {
        //We have no way of knowing what type of IconRead it is until we've read its first node and xr is forward only.
        //This way we can determine type after reading the initial data.
        internal static ExcelConditionalFormattingRule ReadIcons(ExcelAddress address, XmlReader xr, ExcelWorksheet ws)
        {
            //Read base rules
            var priority = int.Parse(xr.GetAttribute("priority"));
            var stopIfTrue = xr.GetAttribute("stopIfTrue") == "1";

            xr.Read();
            var set = xr.GetAttribute("iconSet");

            if (set == null)
            {
                set = "3TrafficLights1";
            }

            //The first char of all iconSet types start with number of their type.
            if (set[0] == '3')
            {
                return new ExcelConditionalFormattingThreeIconSet(address, priority, ws, stopIfTrue, xr);
            }
            else if(set[0] == '4')
            {
                return new ExcelConditionalFormattingFourIconSet(address, priority, ws, stopIfTrue, xr);
            }
            else if(set[0] == '5')
            {
                return new ExcelConditionalFormattingFiveIconSet(address, priority, ws, stopIfTrue, xr);
            }

            throw new NotImplementedException($"{set} is not a known type of IconSet");
        }
    }
}
