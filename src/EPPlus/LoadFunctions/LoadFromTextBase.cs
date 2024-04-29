/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/30/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.LoadFunctions.Params;

namespace OfficeOpenXml.LoadFunctions
{
    internal abstract class LoadFromTextBase<T>
        where T : ExcelTextFileFormat
    {
        protected ExcelWorksheet _worksheet;
        protected ExcelRangeBase _range;
        protected string _text;
        protected T _format;

        public LoadFromTextBase(ExcelRangeBase range, string text, T format)
        {
            _range = range;
            _worksheet = range.Worksheet;
            _text = text;
            _format = format;
        }

        public abstract ExcelRangeBase Load();


        protected string[] SplitLines(string text, string EOL)
        {
            var lines = Regex.Split(text, EOL);
            for (int i = 0; i < lines.Length; i++)
            {
                if (EOL == "\n" && lines[i].EndsWith("\r", StringComparison.OrdinalIgnoreCase)) lines[i] = lines[i].Substring(0, lines[i].Length - 1); //If EOL char is lf and last chart cr then we remove the trailing cr.
                if (EOL == "\r" && lines[i].StartsWith("\n", StringComparison.OrdinalIgnoreCase)) lines[i] = lines[i].Substring(1); //If EOL char is cr and last chart lf then we remove the heading lf.
            }
            return lines;
        }

        protected bool IsEOL(string text, int ix, string eol)
        {
            for (int i = 0; i < eol.Length; i++)
            {
                if (text[ix + i] != eol[i])
                    return false;
            }
            return ix + eol.Length <= text.Length;
        }

        protected object ConvertData(T Format, eDataTypes[] dataType, string v, int col, bool isText)
        {
            if (isText && (dataType == null || dataType.Length < col))
            {
                return string.IsNullOrEmpty(v) ? null : v;
            }
            else
            {
                if(dataType == null || dataType.Length < col)
                    return ConvertData(Format, eDataTypes.Unknown, v, col, isText);
                return ConvertData(Format, dataType[col], v, col, isText);
            }
        }

        protected object ConvertData(T Format, eDataTypes? dataType, string v, int col, bool isText)
        {
            if (isText && dataType == null ) return string.IsNullOrEmpty(v) ? null : v;

            double d;
            DateTime dt;
            if (dataType == null || dataType == eDataTypes.Unknown)
            {
                string v2 = v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v;
                if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out d))
                {
                    if (v2 == v)
                    {
                        return d;
                    }
                    else
                    {
                        return d / 100;
                    }
                }
                if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt))
                {
                    return dt;
                }
                else
                {
                    return string.IsNullOrEmpty(v) ? null : v;
                }
            }
            else
            {
                switch (dataType)
                {
                    case eDataTypes.Number:
                        if (double.TryParse(v, NumberStyles.Any, Format.Culture, out d))
                        {
                            return d;
                        }
                        else
                        {
                            return v;
                        }
                    case eDataTypes.DateTime:
                        if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt))
                        {
                            return dt;
                        }
                        else
                        {
                            return v;
                        }
                    case eDataTypes.Percent:
                        string v2 = v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v;
                        if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out d))
                        {
                            return d / 100;
                        }
                        else
                        {
                            return v;
                        }
                    case eDataTypes.String:
                        return v;
                    default:
                        return string.IsNullOrEmpty(v) ? null : v;

                }
            }
        }

    }
}
