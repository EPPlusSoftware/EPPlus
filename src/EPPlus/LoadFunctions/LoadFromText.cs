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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromText : LoadFromTextBase<ExcelTextFormat>
    {
        public LoadFromText(ExcelRangeBase range, string text, LoadFromTextParams parameters)
            : base(range, text, parameters.Format)
        {
        }

        public override ExcelRangeBase Load()
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
                if (_format.ShouldUseRow != null && _format.ShouldUseRow.Invoke(line) == false)
                {
                    continue;
                }

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
                                if(v.Trim()=="")
                                {
                                    v = "";
                                }
                                else
                                {
                                    throw (new Exception(string.Format("Invalid Text Qualifier in line : {0}", line)));
                                }
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
                                    items.Add(ConvertData(_format, _format.DataTypes, v, col, isText));
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
                        if(string.IsNullOrEmpty(v))
                        {
                            QCount--;
                        }
                        v += new string(_format.TextQualifier, (QCount) / 2);
                    }
                    if (lineQCount % 2 == 1)
                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));

                    items.Add(ConvertData(_format, _format.DataTypes, v, col, isText));
                    if (_format.Transpose)
                    {
                        _worksheet._values.SetValueRow_ValueTranspose(_range._fromRow, _range._fromCol + row, items);
                    }
                    else
                    {
                        _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, items);
                    }

                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }

            if (row <= 0)
            {
                return null;
            }
            if(_format.Transpose)
            {
                return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow +maxCol, _range._fromCol + row - 1];
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol];
        }

        protected string[] GetLines(string text, ExcelTextFormat Format)
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
    }
}
