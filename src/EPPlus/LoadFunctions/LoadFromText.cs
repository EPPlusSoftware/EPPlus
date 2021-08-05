/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromText
    {
        public LoadFromText(ExcelRangeBase range, string text, LoadFromTextParams parameters)
        {
            _range = range;
            _worksheet = range.Worksheet;
            _text = text;
            if(parameters.Format == null)
            {
                _format = new ExcelTextFormat();
            }
            else
            {
                _format = parameters.Format;
            }
        }

        private readonly ExcelWorksheet _worksheet;
        private readonly ExcelRangeBase _range;
        private readonly ExcelTextFormat _format;
        private readonly string _text;

        public ExcelRangeBase Load()
        {
            if (string.IsNullOrEmpty(_text))
            {
                var r = _worksheet.Cells[_range._fromRow, _range._fromCol];
                r.Value = "";
                return r;
            }

            string[] lines;
            if (_format.TextQualifier == 0)
            {
                lines = SplitLines(_text, _format.EOL);
            }
            else
            {
                lines = GetLines(_text, _format);
            }

            int row = 0;
            int col = 0;
            int maxCol = col;
            int lineNo = 1;
            //var values = new List<object>[lines.Length];
            foreach (string line in lines)
            {
                var items = new List<object>();
                //values[row] = items;

                if (lineNo > _format.SkipLinesBeginning && lineNo <= lines.Length - _format.SkipLinesEnd)
                {
                    col = 0;
                    string v = "";
                    bool isText = false, isQualifier = false;
                    int QCount = 0;
                    int lineQCount = 0;
                    foreach (char c in line)
                    {
                        if (_format.TextQualifier != 0 && c == _format.TextQualifier)
                        {
                            if (!isText && v != "")
                            {
                                throw (new Exception(string.Format("Invalid Text Qualifier in line : {0}", line)));
                            }
                            isQualifier = !isQualifier;
                            QCount += 1;
                            lineQCount++;
                            isText = true;
                        }
                        else
                        {
                            if (QCount > 1 && !string.IsNullOrEmpty(v))
                            {
                                v += new string(_format.TextQualifier, QCount / 2);
                            }
                            else if (QCount > 2 && string.IsNullOrEmpty(v))
                            {
                                v += new string(_format.TextQualifier, (QCount - 1) / 2);
                            }

                            if (isQualifier)
                            {
                                v += c;
                            }
                            else
                            {
                                if (c == _format.Delimiter)
                                {
                                    items.Add(ConvertData(_format, v, col, isText));
                                    v = "";
                                    isText = false;
                                    col++;
                                }
                                else
                                {
                                    if (QCount % 2 == 1)
                                    {
                                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));
                                    }
                                    v += c;
                                }
                            }
                            QCount = 0;
                        }
                    }
                    if (QCount > 1)
                    {
                        v += new string(_format.TextQualifier, QCount / 2);
                    }
                    if (lineQCount % 2 == 1)
                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));
                    items.Add(ConvertData(_format, v, col, isText));

                    _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, items);

                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }

            if (row <= 0)
            {
                return null;
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol];
        }

        private string[] SplitLines(string text, string EOL)
        {
            var lines = Regex.Split(text, EOL);
            for (int i = 0; i < lines.Length; i++)
            {
                if (EOL == "\n" && lines[i].EndsWith("\r", StringComparison.OrdinalIgnoreCase)) lines[i] = lines[i].Substring(0, lines[i].Length - 1); //If EOL char is lf and last chart cr then we remove the trailing cr.
                if (EOL == "\r" && lines[i].StartsWith("\n", StringComparison.OrdinalIgnoreCase)) lines[i] = lines[i].Substring(1); //If EOL char is cr and last chart lf then we remove the heading lf.
            }
            return lines;
        }

        private string[] GetLines(string text, ExcelTextFormat Format)
        {
            if (Format.EOL == null || Format.EOL.Length == 0) return new string[] { text };
            var eol = Format.EOL;
            var list = new List<string>();
            var inTQ = false;
            var prevLineStart = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == Format.TextQualifier)
                {
                    inTQ = !inTQ;
                }
                else if (!inTQ)
                {
                    if (IsEOL(text, i, eol))
                    {
                        var s = text.Substring(prevLineStart, i - prevLineStart);
                        if (eol == "\n" && s.EndsWith("\r", StringComparison.OrdinalIgnoreCase)) s = s.Substring(0, s.Length - 1); //If EOL char is lf and last chart cr then we remove the trailing cr.
                        if (eol == "\r" && s.StartsWith("\n", StringComparison.OrdinalIgnoreCase)) s = s.Substring(1); //If EOL char is cr and last chart lf then we remove the heading lf.
                        list.Add(s);
                        i += eol.Length - 1;
                        prevLineStart = i + 1;
                    }
                }
            }

            if (inTQ)
            {
                throw (new ArgumentException(string.Format("Text delimiter is not closed in line : {0}", list.Count)));
            }

            //if (text.Length >= Format.EOL.Length && IsEOL(text, text.Length-2, Format.EOL))
            //{
            //    //list.Add(text.Substring(prevLineStart- Format.EOL.Length, Format.EOL.Length));
            //    list.Add("");
            //}
            //else
            //{
                list.Add(text.Substring(prevLineStart));
            //}
            return list.ToArray();
        }
        private bool IsEOL(string text, int ix, string eol)
        {
            for (int i = 0; i < eol.Length; i++)
            {
                if (text[ix + i] != eol[i])
                    return false;
            }
            return ix + eol.Length <= text.Length;
        }

        private object ConvertData(ExcelTextFormat Format, string v, int col, bool isText)
        {
            if (isText && (Format.DataTypes == null || Format.DataTypes.Length < col)) return string.IsNullOrEmpty(v) ? null : v;

            double d;
            DateTime dt;
            if (Format.DataTypes == null || Format.DataTypes.Length <= col || Format.DataTypes[col] == eDataTypes.Unknown)
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
                switch (Format.DataTypes[col])
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
