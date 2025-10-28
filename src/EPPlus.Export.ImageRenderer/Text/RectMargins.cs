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
using OfficeOpenXml.Drawing;

namespace EPPlusImageRenderer.Text
{
    internal class RectMargins : RectBase
    {
        internal double MarginLeft { get; set; }
        internal double MarginTop { get; set; }
        internal double MarginRight { get; set; }
        internal double MarginBottom { get; set; }

        internal RectMargins(): base() { }

        internal RectMargins(double l, double t, double r, double b) : base(l, t, r, b) 
        { 
        }

        internal double GetInnerLeft()
        {
            return Left + MarginLeft;
        }

        internal double GetInnerTop()
        {
            return Top + MarginTop;
        }

        internal double GetInnerBottom()
        {
            return Bottom - MarginBottom;
        }

        internal double GetInnerRight()
        {
            return Right - MarginRight;
        }

        public RectBase GetInnerRect()
        {
            var textArea = new RectBase();
            textArea.Left = GetInnerLeft();
            textArea.Top = GetInnerTop();

            textArea.Bottom = GetInnerBottom();
            textArea.Right = GetInnerRight();

            return textArea;
        }

        public double GetInnerWidth()
        {
            return GetInnerRight() - GetInnerLeft();
        }

        public double GetInnerHeight()
        {
            return GetInnerBottom() - GetInnerTop();
        }
    }
}
