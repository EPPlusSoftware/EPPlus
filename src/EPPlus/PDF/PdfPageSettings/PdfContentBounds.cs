using OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfPageSettings
{
    internal class PdfContentBounds : PdfRect
    {
        public PdfContentBounds(PdfMargins margins, PdfPageSize pageSize)
        {
            X = margins.LeftPoints;
            Y = margins.BottomPoints;
            Width = pageSize.WidthPoints - margins.LeftPoints - margins.RightPoints;
            Height = pageSize.HeightPoints - margins.TopPoints - margins.BottomPoints;
        }
    }
}
