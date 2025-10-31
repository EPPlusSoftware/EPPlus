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
using EPPlus.Export.Pdf.Pdfhelpers;

namespace EPPlus.Export.Pdf.PdfSettings.PdfPageSizes
{
    public class PdfMargins
    {
        public double Header { get; set; } = 7.6d;
        public double Footer { get; set; } = 7.6d;
        public double Top { get; set; } = 19.1d;
        public double Left { get; set; } = 17.8d;
        public double Right { get; set; } = 17.8d;
        public double Bottom { get; set; } = 19.1d;
        public double HeaderPu { get; private set; }
        public double FooterPu { get; private set; }
        public double TopPu { get; private set; }
        public double LeftPu { get; private set; }
        public double RightPu { get; private set; }
        public double BottomPu { get; private set; }

        public PdfMargins()
        {
            TopPu = PdfUnits.MmToPoints(Top);
            LeftPu = PdfUnits.MmToPoints(Left);
            RightPu = PdfUnits.MmToPoints(Right);
            BottomPu = PdfUnits.MmToPoints(Bottom);
            HeaderPu = PdfUnits.MmToPoints(Header);
            FooterPu = PdfUnits.MmToPoints(Footer);
        }

        public PdfMargins(double Top, double Left, double Right, double Bottom, double Header, double Footer)
        {
            this.Top = Top;
            this.Left = Left;
            this.Right = Right;
            this.Bottom = Bottom;
            this.Header = Header;
            this.Footer = Footer;
            TopPu = PdfUnits.MmToPoints(Top);
            LeftPu = PdfUnits.MmToPoints(Left);
            RightPu = PdfUnits.MmToPoints(Right);
            BottomPu = PdfUnits.MmToPoints(Bottom);
            HeaderPu = PdfUnits.MmToPoints(Header);
            FooterPu = PdfUnits.MmToPoints(Footer);
        }

        public static PdfMargins Normal => new PdfMargins();
        public static PdfMargins Wide => new PdfMargins(25.4d, 25.4d, 25.4d, 25.4d, 12.7d, 12.7d);
        public static PdfMargins Narrow => new PdfMargins(19.1d, 6.4d, 6.4d, 19.1d, 7.6d, 7.6d);
    }
}
