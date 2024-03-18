/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Style;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal static class HtmlRichText
    {
        internal static void GetRichTextStyle(ExcelRichText rt, StringBuilder sb)
        {
            if(rt.Bold)
            {
                sb.Append("font-weight:bolder;");
            }
            if (rt.Italic)
            {
                sb.Append("font-style:italic;");
            }
            if(rt.UnderLine)
            {
                sb.Append("text-decoration:underline solid;");
            }
            if (rt.Strike)
            {
                sb.Append("text-decoration:line-through solid;");
            }
            if(rt.Size > 0)
            {
                sb.Append($"font-size:{rt.Size.ToString("g", CultureInfo.InvariantCulture)}pt;");
            }
            if (string.IsNullOrEmpty(rt.FontName)==false)
            {
                sb.Append($"font-family:{rt.FontName};");
            }
            if(rt.Color.IsEmpty==false)
            {
                sb.Append("color:#" + rt.Color.ToArgb().ToString("x8").Substring(2));
            }
        }
    }
}
