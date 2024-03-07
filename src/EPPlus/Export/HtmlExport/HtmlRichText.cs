using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal static class HtmlRichText
    {
        internal static void GetRichTextStyle(ExcelRichText rt, StringBuilder sb)
        {
            if(rt.Bold != null)
            {
                sb.Append("font-weight:bolder;");
            }
            if (rt.Italic != null)
            {
                sb.Append("font-style:italic;");
            }
            if(rt.UnderLine)
            {
                sb.Append("text-decoration:underline solid;");
            }
            if (rt.Strike != null)
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

        //internal static void GetRichTextStyle(ExcelRichText rt, StringBuilder sb)
        //{
        //    if (rt.Bold)
        //    {
        //        sb.Append("font-weight:bolder;");
        //    }
        //    if (rt.Italic)
        //    {
        //        sb.Append("font-style:italic;");
        //    }
        //    if (rt.UnderLine)
        //    {
        //        sb.Append("text-decoration:underline solid;");
        //    }
        //    if (rt.Strike)
        //    {
        //        sb.Append("text-decoration:line-through solid;");
        //    }
        //    if (rt.Size > 0)
        //    {
        //        sb.Append($"font-size:{rt.Size.ToString("g", CultureInfo.InvariantCulture)}pt;");
        //    }
        //    if (string.IsNullOrEmpty(rt.FontName) == false)
        //    {
        //        sb.Append($"font-family:{rt.FontName};");
        //    }
        //    if (rt.Color.IsEmpty == false)
        //    {
        //        sb.Append("color:#" + rt.Color.ToArgb().ToString("x8").Substring(2));
        //    }
        //}
    }
}
