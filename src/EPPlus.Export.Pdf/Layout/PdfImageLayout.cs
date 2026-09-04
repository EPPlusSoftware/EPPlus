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
using EPPlus.Graphics;
using System.Diagnostics;

namespace EPPlus.Export.Pdf.Layout
{
    [DebuggerDisplay("Image: {Name}")]
    internal class PdfImageLayout : Transform
    {
        public byte[] ImageBytes;

        public bool IsHeaderFooter;

        public PdfImageLayout(double x, double y, double width, double height)
            : base(x, y - height, width, height)
        {
            Z = 5; // paint above cell fills, text and borders
        }
    }
}