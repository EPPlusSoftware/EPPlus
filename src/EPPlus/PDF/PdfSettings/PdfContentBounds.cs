using OfficeOpenXml.PDF.PdfSettings.PdfPageSizes;

namespace OfficeOpenXml.PDF.PdfSettings
{
    internal class PdfContentBounds : PdfRect
    {
        public double HeaderY;
        public double FooterY;
        public double CenterHeaderX;
        public double RightHeaderX;
        public double CenterFooterX;
        public double RightFooterX;

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
