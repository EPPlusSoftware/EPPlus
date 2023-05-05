using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
