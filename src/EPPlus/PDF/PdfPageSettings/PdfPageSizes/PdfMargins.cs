using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes
{
    public class PdfMargins
    {
        public float Header { get; set; } = 8f;
        public float Footer { get; set; } = 8f;
        public float Top { get; set; } = 19f;
        public float Left { get; set; } = 18f;
        public float Right { get; set; } = 18f;
        public float Bottom { get; set; } = 19f;

        public int HeaderPoints { get; private set; }
        public int FooterPoints { get; private set; }
        public int TopPoints { get; private set; }
        public int LeftPoints { get; private set; }
        public int RightPoints { get; private set; }
        public int BottomPoints { get; private set; }

        public PdfMargins()
        {
            TopPoints = PdfUnits.MmToPointsRounded(Top);
            LeftPoints = PdfUnits.MmToPointsRounded(Left);
            RightPoints = PdfUnits.MmToPointsRounded(Right);
            BottomPoints = PdfUnits.MmToPointsRounded(Bottom);
        }

        public PdfMargins(float Top, float Left, float Right, float Bottom)
        {
            this.Top = Top;
            this.Left = Left;
            this.Right = Right;
            this.Bottom = Bottom;

            TopPoints = PdfUnits.MmToPointsRounded(Top);
            LeftPoints = PdfUnits.MmToPointsRounded(Left);
            RightPoints = PdfUnits.MmToPointsRounded(Right);
            BottomPoints = PdfUnits.MmToPointsRounded(Bottom);
        }

    }
}
