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
using System.Drawing;
using EPPlus.Graphics;

namespace EPPlus.Export.Pdf.Layout
{
    internal class PdfDataBarLayout : Transform
    {
        public double Percentage;
        public Color FillColor;
        public bool IsNegative;
        public double AxisFraction;
        public System.Drawing.Color? BorderColor;
        public System.Drawing.Color? AxisColor;
        public bool Gradient;
        public bool IsRightToLeft;

        public PdfDataBarLayout(double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y - height, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 1;
        }
    }
}