/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Graphics;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.Layout
{
    internal struct PdfNamedDest
    {
        public string Name;
        public string Sheet;
        public int Row;
        public int Col;
        public string CellText;
    }

    internal class PdfCatalogLayout : Transform
    {
        internal List<PdfNamedDest> NamedDestinations = new List<PdfNamedDest>();

        public PdfCatalogLayout(double x, double y, double width, double height)
            : base(x, y, width, height)
        { }
    }
}