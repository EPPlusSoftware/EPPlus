using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class RangeCopyStylesHelper
    {
        private readonly ExcelRangeBase _sourceRange;
        private readonly ExcelRangeBase _destinationRange;
        internal RangeCopyStylesHelper(ExcelRangeBase sourceRange, ExcelRangeBase destinationRange)
        {
            _sourceRange = sourceRange;
            _destinationRange = destinationRange;
        }
        internal void CopyStyles()
        {
            var styleCashe = new Dictionary<int, int>();
            var wsSource = _sourceRange.Worksheet;
            var wsDest= _destinationRange.Worksheet;
            var sameWorkbook = wsSource.Workbook == wsDest.Workbook; 
            var sc = _sourceRange._fromCol;
            for(int dc=_destinationRange._fromCol; dc <= _destinationRange._toCol; dc++)
            {
                var sr = _sourceRange._fromRow;
                for (int dr = _destinationRange._fromRow; dr <= _destinationRange._toRow; dr++)
                {
                    int styleId = GetStyleId(wsSource, sc, sr);
                    if (!sameWorkbook)
                    {
                        if (styleCashe.ContainsKey(styleId))
                        {
                            styleId = styleCashe[styleId];
                        }
                        else
                        {
                            var sourceStyleId = styleId;
                            styleId = wsDest.Workbook.Styles.CloneStyle(wsSource.Workbook.Styles, styleId);
                            styleCashe.Add(sourceStyleId, styleId);
                        }
                    }
                    _destinationRange.Worksheet.SetStyleInner(dr, dc, styleId);

                    if (sr < _sourceRange._toRow) sr++;
                }
                if (sc < _sourceRange._toCol) sc++;
            }
        }

        private static int GetStyleId(ExcelWorksheet wsSource, int sc, int sr)
        {
            var styleId = wsSource.GetStyleInner(sr, sc);
            if (styleId == 0)
            {
                styleId = wsSource.GetStyleInner(sr, 0);
                if (styleId == 0)
                {
                    styleId = wsSource.GetStyleInner(0, sc);
                }
            }

            return styleId;
        }
    }
}
