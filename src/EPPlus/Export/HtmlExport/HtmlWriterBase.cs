/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal abstract class HtmlWriterBase
    {
        internal HtmlWriterBase(Stream stream, Encoding encoding)
        {
            _stream = stream;
            _writer = new StreamWriter(stream, encoding);
            
        }
        public HtmlWriterBase(StreamWriter writer)
        {
            _stream = writer.BaseStream;
            _writer = writer;
        }
        protected readonly Stream _stream;
        protected readonly StreamWriter _writer;

        protected const string IndentWhiteSpace = "  ";
        protected bool _newLine;

        protected Dictionary<string, int> _styleCache = new Dictionary<string, int>();

        internal int Indent { get; set; }

        protected internal static bool HasStyle(ExcelXfs xfs)
        {
            return xfs.FontId > 0 ||
                   xfs.FillId > 0 ||
                   xfs.BorderId > 0 ||
                   xfs.HorizontalAlignment != ExcelHorizontalAlignment.General ||
                   xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom ||
                   xfs.TextRotation != 0 ||
                   xfs.Indent > 0 ||
                   xfs.WrapText;
        }
        protected internal static string GetStyleKey(ExcelXfs xfs)
        {
            var fbfKey = (ulong)(xfs.FontId << 32 | xfs.BorderId << 16 | xfs.FillId);
            return fbfKey.ToString() + "|" + ((int)xfs.HorizontalAlignment).ToString() + "|" + ((int)xfs.VerticalAlignment).ToString() + "|" + xfs.Indent.ToString() + "|" + xfs.TextRotation.ToString() + "|" + (xfs.WrapText ? "1" : "0");
        }

        protected static string WriteBorderItemLine(ExcelBorderStyle style, string suffix)
        {
            var lineStyle = $"border-{suffix}:";
            switch (style)
            {
                case ExcelBorderStyle.Hair:
                    lineStyle += "1px solid";
                    break;
                case ExcelBorderStyle.Thin:
                    lineStyle += $"thin solid";
                    break;
                case ExcelBorderStyle.Medium:
                    lineStyle += $"medium solid";
                    break;
                case ExcelBorderStyle.Thick:
                    lineStyle += $"thick solid";
                    break;
                case ExcelBorderStyle.Double:
                    lineStyle += $"double";
                    break;
                case ExcelBorderStyle.Dotted:
                    lineStyle += $"dotted";
                    break;
                case ExcelBorderStyle.Dashed:
                case ExcelBorderStyle.DashDot:
                case ExcelBorderStyle.DashDotDot:
                    lineStyle += $"dashed";
                    break;
                case ExcelBorderStyle.MediumDashed:
                case ExcelBorderStyle.MediumDashDot:
                case ExcelBorderStyle.MediumDashDotDot:
                    lineStyle += $"medium dashed";
                    break;
            }
            return lineStyle;
        }
        protected static string GetVerticalAlignment(ExcelXfs xfs)
        {
            switch (xfs.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    return "top";
                case ExcelVerticalAlignment.Center:
                    return "middle";
                case ExcelVerticalAlignment.Bottom:
                    return "bottom";
            }

            return "";
        }

        protected static string GetHorizontalAlignment(ExcelXfs xfs)
        {
            switch (xfs.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Right:
                    return "right";
                case ExcelHorizontalAlignment.Center:
                case ExcelHorizontalAlignment.CenterContinuous:
                    return "center";
                case ExcelHorizontalAlignment.Left:
                    return "left";
            }

            return "";
        }
        public void WriteLine()
        {
            _newLine = true;
            _writer.WriteLine();
        }

        public void Write(string text)
        {
            _writer.Write(text);
        }

        internal protected void WriteIndent()
        {
            for (var x = 0; x < Indent; x++)
            {
                _writer.Write(IndentWhiteSpace);
            }
        }
        internal void ApplyFormat(bool minify)
        {
            if (minify == false)
            {
                WriteLine();
            }
        }

        internal void ApplyFormatIncreaseIndent(bool minify)
        {
            if (minify == false)
            {
                WriteLine();
                Indent++;
            }
        }

        internal void ApplyFormatDecreaseIndent(bool minify)
        {
            if (minify == false)
            {
                WriteLine();
                Indent--;
            }
        }
        internal void WriteClass(string value, bool minify)
        {
            if (minify)
            {
                _writer.Write(value);
            }
            else
            {
                _writer.WriteLine(value);
                Indent = 1;
            }
        }
        internal void WriteClassEnd(bool minify)
        {
            if (minify)
            {
                _writer.Write("}");
            }
            else
            {
                _writer.WriteLine("}");
                Indent = 0;
            }
        }
        internal void WriteCssItem(string value, bool minify)
        {
            if (minify)
            {
                _writer.Write(value);
            }
            else
            {
                WriteIndent();
                _writer.WriteLine(value);
            }
        }
    }
}