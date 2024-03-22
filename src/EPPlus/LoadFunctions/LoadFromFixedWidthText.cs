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

            if(_format.ReadStartPosition == FixedWidthReadType.Widths)
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
            var maxCol = col;
            var row = 0;
            var lineNo = 1;
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
                    if(line.Length < _format.LineLength)
                    {
                        continue;
                    }
                    var items = new List<object>();
                    var isText = false;
                    int readLength = 0;
                    col = 0;
                    for (int i = 0; i < _format.ColumnFormat.Count; i++)
                    {
                        string content;
                        if (i == 0)
                        {
                            content = line.Substring(0, _format.ColumnFormat[i].Length);
                            readLength += _format.ColumnFormat[i].Length;
                        }
                        else
                        {
                            var v = line.Length;
                            if (readLength + _format.ColumnFormat[i].Length >= v)
                            {
                                content = line.Substring(readLength + 1);
                            }
                            else
                            {
                                content = line.Substring(readLength, _format.ColumnFormat[i].Length);
                                readLength += _format.ColumnFormat[i].Length;
                            }
                        }
                        content = content.Trim();
                        if (_format.ColumnFormat[i].UseColumn)
                        {
                            items.Add(ConvertData(_format, _format.ColumnFormat[i].DataType, content.Trim(), col, isText));
                            col++;
                        }
                    }
                    _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, items);
                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol ];
        }

        private ExcelRangeBase LoadPositions()
        {
            string[] lines;
            lines = SplitLines(_text, _format.EOL);
            var col = 0;
            var maxCol = col;
            var row = 0;
            var lineNo = 1;
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
                    if(line.Length <= _format.LineLength)
                    {
                        continue;
                    }
                    var items = new List<object>();
                    var isText = false;
                    col = 0;
                    for (int i = 0; i < _format.ColumnFormat[i].Position; i++)
                    {
                        string content;
                        if (i == _format.ColumnFormat[i].Position - 1)
                        {
                            content = line.Substring(_format.ColumnFormat[i].Position);
                        }
                        else
                        {
                            var readLength = _format.ColumnFormat[i + 1].Position - _format.ColumnFormat[i].Position;
                            content = line.Substring(_format.ColumnFormat[i].Position, readLength);
                        }
                        content = content.Trim();
                        if (_format.ColumnFormat[i].UseColumn)
                        {
                            items.Add(ConvertData(_format, _format.ColumnFormat[i].DataType, content.Trim(), col, isText));
                            col++;
                        }
                    }
                    _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, items);
                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol - 1];
        }

    }
}
