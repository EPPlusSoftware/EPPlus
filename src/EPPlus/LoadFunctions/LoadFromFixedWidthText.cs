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
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromFixedWidthText : LoadFromTextBase<ExcelTextFormatFixedWidth>
    {

        public LoadFromFixedWidthText(ExcelRangeBase range, string text, ExcelTextFormatFixedWidth Format)
            : base(range, text, Format)
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

            if (_format.ReadType == FixedWidthReadType.Length)
            {
                return LoadWidths();
            }
            else
            {
                return LoadPositions();
            }
        }

        private ExcelRangeBase LoadWidths()
        {
            string[] lines;
            lines = SplitLines(_text, _format.EOL);
            var col = 0;
            var maxCol = 1;
            var row = 0;
            var lineNo = 1;

            var columnNames = new List<object>();
            for (int i = 0; i < _format.Columns.Count; i++)
            {
                if (!string.IsNullOrEmpty(_format.Columns[i].Name) && _format.Columns[i].UseColumn && row == 0)
                {
                    columnNames.Add(_format.Columns[i].Name);
                    col++;
                }
            }
            if (columnNames.Count > 0)
            {
                _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, columnNames);
                if (col > maxCol) maxCol = col;
                row++;
            }

            foreach (string line in lines)
            {
                if (lineNo > _format.SkipLinesBeginning && lineNo <= lines.Length - _format.SkipLinesEnd)
                {
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }

                    if (_format.ShouldUseRow != null && _format.ShouldUseRow.Invoke(line) == false)
                    {
                        continue;
                    }

                    var items = new List<object>();
                    var isText = false;
                    int readLength = 0;
                    col = 0;
                    bool lineread = false;

                    if ( line.Length < _format.LineLength && _format.FormatErrorStrategy == FixedWidthFormatErrorStrategy.ThrowError )
                    {
                        throw new FormatException("Line was " + line.Length + ", Expected length of " + _format.LineLength + " at Line " + lineNo );
                    }

                    for (int i = 0; i < _format.Columns.Count; i++)
                    {
                        string content;
                        if (lineread)
                        {
                            continue;
                        }
                        if (_format.Columns[i].Length <= 0 || readLength + _format.Columns[i].Length >= line.Length)
                        {
                            content = line.Substring(readLength);
                            lineread = true;
                        }
                        else
                        {
                            content = line.Substring(readLength, _format.Columns[i].Length);
                            readLength += _format.Columns[i].Length;
                        }
                        content = content.Trim(_format.PaddingCharacter);
                        if (_format.Columns[i].UseColumn)
                        {
                            items.Add(ConvertData(_format, _format.Columns[i].DataType, content.Trim(), col, isText));
                            col++;
                        }
                    }
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
            if (_format.Transpose)
            {
                return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + maxCol - 1, _range._fromCol + row - 1];
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol - 1];
        }

        private ExcelRangeBase LoadPositions()
        {
            string[] lines;
            lines = SplitLines(_text, _format.EOL);
            var col = 0;
            var maxCol = 1;
            var row = 0;
            var lineNo = 1;

            var columnNames = new List<object>();
            for (int i = 0; i < _format.Columns.Count; i++)
            {
                if (!string.IsNullOrEmpty(_format.Columns[i].Name) && _format.Columns[i].UseColumn && row == 0)
                {
                    columnNames.Add(_format.Columns[i].Name);
                    col++;
                }
            }
            if (columnNames.Count > 0)
            {
                _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, columnNames);
                if (col > maxCol) maxCol = col;
                row++;
            }


            foreach (string line in lines)
            {
                if (lineNo > _format.SkipLinesBeginning && lineNo <= lines.Length - _format.SkipLinesEnd)
                {
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }
                    if (_format.ShouldUseRow != null && _format.ShouldUseRow.Invoke(line) == false)
                    {
                        continue;
                    }
                    if (line.Length < _format.LineLength && _format.FormatErrorStrategy == FixedWidthFormatErrorStrategy.ThrowError)
                    {
                        throw new FormatException("Line was " + line.Length + ", Expected length of " + _format.LineLength);
                    }
                    var isText = false;
                    var items = new List<object>();
                    col = 0;
                    for (int i = 0; i < _format.Columns.Count; i++)
                    {
                        if (line.Length < _format.Columns[i].Position)
                        {
                            continue;
                        }
                        string content;
                        if (i == _format.Columns.Count - 1)
                        {
                            if (_format.Columns[i].Length > 0)
                            {
                                content = line.Substring(_format.Columns[i].Position, _format.Columns[i].Length);
                            }
                            else
                            {
                                content = line.Substring(_format.Columns[i].Position);
                            }
                        }
                        else if( i == 0 )
                        {
                            var readLength = _format.Columns[i + 1].Position - _format.Columns[i].Position;
                            content = line.Substring(_format.Columns[i].Position, readLength);
                        }
                        else
                        {
                            var readLength = _format.Columns[i + 1].Position - _format.Columns[i].Position;
                            if ((readLength > _format.Columns[i].Position && _format.FormatErrorStrategy == FixedWidthFormatErrorStrategy.Truncate))
                            {
                                content = line.Substring(_format.Columns[i].Position);
                            }
                            else
                            {
                                content = line.Substring(_format.Columns[i].Position, readLength);
                            }
                        }
                        content = content.Trim(_format.PaddingCharacter);

                        if (_format.Columns[i].UseColumn)
                        {
                            items.Add(ConvertData(_format, _format.Columns[i].DataType, content.Trim(), col, isText));
                            col++;
                        }
                    }
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
            if (_format.Transpose)
            {
                return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + maxCol - 1, _range._fromCol + row - 1];
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol - 1];
        }

    }
}
