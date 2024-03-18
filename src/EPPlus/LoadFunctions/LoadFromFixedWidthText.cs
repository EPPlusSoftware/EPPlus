using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromFixedWidthText : LoadFromText
    {
        public LoadFromFixedWidthText(ExcelRangeBase range, string text, LoadFromTextParams parameters, params int[] columnLengths) 
            : base(range, text, parameters)
        {
            _columnLengths = columnLengths;
        }

        private int[] _columnLengths;

        public override ExcelRangeBase Load()
        {
            if (string.IsNullOrEmpty(_text))
            {
                var r = _worksheet.Cells[_range._fromRow, _range._fromCol];
                r.Value = "";
                return r;
            }

            string[] lines;
            lines = SplitLines(_text, _format.EOL);
            var col = 0;
            var maxCol = col;
            var row = 0;
            foreach (string line in lines)
            {
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }
                var items = new List<object>();
                var isText = false;
                int readLength = 0;
                for (int i = 0; i < _columnLengths.Length; i++)
                {
                    string content;
                    if (i == 0)
                    {
                        content = line.Substring(0, _columnLengths[i]);
                        readLength += _columnLengths[i];
                    }
                    else
                    {
                        var v = line.Length;
                        if (readLength + _columnLengths[i] >= v)
                        {
                            content = line.Substring(readLength + 1);
                        }
                        else
                        {
                            content = line.Substring(readLength + 1, _columnLengths[i]);
                            readLength += _columnLengths[i];
                        }
                    }
                    content = content.Trim();
                    items.Add(ConvertData(_format, content.Trim(), col, isText));
                    col++;
                }
                _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, items);
                row++;
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol];

            //private ExcelRangeBase LoadfixedWidthFile(string file, ExcelRangeBase startCell, int NoCols, params int[] widths)
            //{
            //    if(NoCols == widths.Length)
            //    {
            //        var currentcell = startCell;
            //        for(int i = 0; i < NoCols; i++)
            //        {
            //            string s = file.read(widths[i]);
            //            currentcell.value = s;
            //            currentcell = currentcell.NextCell();
            //        }
            //    }
            //    else
            //    {
            //        throw new InvalidOperationException("NoCols and widths mismatch, NoCols Needs to be the same as widths length");
            //    }
            //}
        }
    }
}
