using OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes;

namespace OfficeOpenXml.PDF.PdfPageSettings
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
            Width = pageSize.WidthPu - margins.LeftPu - margins.RightPu;
            X = margins.LeftPu;
            Left = X;
            Right = X + Width;

            Y = margins.BottomPu;
            Top = pageSize.HeightPu - margins.TopPu;
            Bottom = Y;
            Height = Top - Bottom;


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
