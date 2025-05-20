using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.PDF.Pdfhelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes
{
    public class PdfMargins
    {
        public double Header { get; set; } = 7.6f;
        public double Footer { get; set; } = 7.6f;
        public double Top { get; set; } = 19.1f;
        public double Left { get; set; } = 17.8f;
        public double Right { get; set; } = 17.8f;
        public double Bottom { get; set; } = 19.1f;

        public double HeaderPoints { get; private set; }
        public double FooterPoints { get; private set; }
        public double TopPoints { get; private set; }
        public double LeftPoints { get; private set; }
        public double RightPoints { get; private set; }
        public double BottomPoints { get; private set; }

        public PdfMargins()
        {
            TopPoints = PdfUnits.MmToPoints(Top);
            LeftPoints = PdfUnits.MmToPoints(Left);
            RightPoints = PdfUnits.MmToPoints(Right);
            BottomPoints = PdfUnits.MmToPoints(Bottom);
            HeaderPoints = PdfUnits.MmToPoints(Header);
            FooterPoints = PdfUnits.MmToPoints(Footer);
        }

        public PdfMargins(double Top, double Left, double Right, double Bottom)
        {
            this.Top = Top;
            this.Left = Left;
            this.Right = Right;
            this.Bottom = Bottom;

            TopPoints = PdfUnits.MmToPoints(Top);
            LeftPoints = PdfUnits.MmToPoints(Left);
            RightPoints = PdfUnits.MmToPoints(Right);
            BottomPoints = PdfUnits.MmToPoints(Bottom);
        }

    }
}
