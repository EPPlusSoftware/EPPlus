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
using OfficeOpenXml;
using EPPlus.Graphics;

namespace EPPlus.Export.Pdf.PdfLayout
{
    /// <summary>
    /// Represents the pages object in the pdf document.
    /// </summary>
    internal class PdfPagesLayout : Transform
    {
        internal ExcelRangeBase Range;

        internal PdfPagesLayout(double x, double y, double height, double width)
            : base(x, y, height, width)
        {
        }

    }
}
