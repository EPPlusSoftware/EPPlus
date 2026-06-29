/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal struct RowHeight
    {
        public double Height;
        public bool UsesDefaultValue;
    }

    internal struct PdfRange
    {
        public List<RowHeight> RowHeights = new List<RowHeight>();
        public List<double> ColWidths = new List<double>();
        public double TotalHeight;
        public double TotalWidth;
        public double AdditionalHeight;
        public double AdditionalWidth;
        public double PrintTitleHeight;
        public double PrintTitleWidth;
        public int PrintTitleRowTo = -1;
        public int PrintTitleColTo = -1;

        public ExcelRangeBase Range { get; set; }
        public bool ExtendColumns { get; set; }
        public PdfCellCollection Map { get; set; }

        public PdfRange(ExcelRangeBase range, bool extendColumns)
        {
            Range = range;
            ExtendColumns = extendColumns;
        }
    }
}
