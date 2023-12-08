/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/7/2023         EPPlus Software AB       EPPlus 7.0.4
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class ColumnsHeaderReader
    {
        public static string GetHeader(MemberInfo member, string headerPrefix)
        {
            var header = member.Name;
            var colAttr = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
            if(colAttr != null && !string.IsNullOrEmpty(colAttr.Header))
            {

                header = colAttr.Header;
            }
            if (!string.IsNullOrEmpty(headerPrefix))
            {
                header = $"{headerPrefix} {header}";
            }
            return header;
        }

        public static string GetAggregatedHeaderPrefix(string aggregatedPrefix, EpplusNestedTableColumnAttribute attr)
        {
            var hPrefix = attr.HeaderPrefix;
            if (!string.IsNullOrEmpty(aggregatedPrefix) && !string.IsNullOrEmpty(hPrefix))
            {
                hPrefix = $"{aggregatedPrefix} {hPrefix}";
            }
            else if (!string.IsNullOrEmpty(aggregatedPrefix))
            {
                hPrefix = aggregatedPrefix;
            }
            return hPrefix;
        }
    }

}
