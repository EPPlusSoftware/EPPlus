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
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics;
using System.Collections.Generic;
using System.Diagnostics;

namespace EPPlus.Export.Pdf.Layout
{
    [DebuggerDisplay("{Name}")]
    internal class PdfPageLayout : Transform
    {
        internal List<GridLine> GridLines = new List<GridLine>();
        internal List<GridLine> BorderLines = new List<GridLine>();
        public List<GridLine> PrintTitleGridLines = new List<GridLine>();
        public double HeadingWidth;
        public double HeadingHeight;
        public double PrintTitleWidth;
        public double PrintTitleHeight;
        public bool isCommentsPage = false;
        /// <summary>
        /// The settings of the worksheet this page belongs to.
        /// Set in PdfLayout.GetCatalog, read in ExcelPdf when writing the page.
        /// </summary>
        internal PdfPageSettings Settings;
        public PdfPageLayout(double x, double y, double width, double height)
            : base(x, y, width, height)
        { }
    }
}