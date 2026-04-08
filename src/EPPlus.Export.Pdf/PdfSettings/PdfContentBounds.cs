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
using EPPlus.Export.Pdf.PdfSettings.PdfPageSizes;
using EPPlus.Graphics;


namespace EPPlus.Export.Pdf.PdfSettings
{
    internal class PdfContentBounds : Rect
    {
        public double HeaderY;
        public double FooterY;
        public double CenterHeaderX;
        public double RightHeaderX;
        public double CenterFooterX;
        public double RightFooterX;

        private double _top;
        private double _bottom;
        private double _left;
        private double _right;
        private double _height;
        private double _width;

        internal override double Top
        {
            get { return _top; }
            set { _top = value; }
        }

        internal override double Bottom
        {
            get { return _bottom; }
            set { _bottom = value; }
        }

        internal override double Left
        {
            get { return _left; }
            set { _left = value; }
        }

        internal override double Right
        {
            get { return _right; }
            set { _right = value; }
        }

        internal override double Width
        {
            get { return _width; }
            set { _width = value; }
        }

        internal override double Height
        {
            get { return _height; }
            set { _height = value; }
        }

        public PdfContentBounds(PdfMargins margins, PdfPageSize pageSize)
        {
            CalculateBounds(margins, pageSize);
        }

        internal void CalculateBounds(PdfMargins margins, PdfPageSize pageSize)
        {
            //Content bounds rectangle
            Width = pageSize.WidthPu - margins.LeftPu - margins.RightPu;
            X = margins.LeftPu;
            Left = X;
            Right = X + Width;
            Y = margins.BottomPu;
            Top = pageSize.HeightPu - margins.TopPu;
            Bottom = Y;
            Height = Top - Bottom;
            //Header Footer
            var hx = Width / 3d;
            FooterY = margins.FooterPu;
            CenterFooterX = Left + hx;
            RightFooterX = Left + hx * 2;
            HeaderY = pageSize.HeightPu - margins.HeaderPu;
            CenterHeaderX = Left + hx;
            RightHeaderX = Left + hx * 2;
        }
    }
}
