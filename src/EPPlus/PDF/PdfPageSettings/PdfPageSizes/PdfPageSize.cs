using OfficeOpenXml.PDF.Pdfhelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes
{
    public class PdfPageSize
    {
        public double Width { get; }
        public double Height { get; }

        public double WidthPoints { get; }
        public double HeightPoints { get; }


        public PdfPageSize(double width, double height)
        {
            Width = width;
            Height = height;
            WidthPoints = Math.Round( PdfUnits.MmToPoints(width));
            HeightPoints = Math.Round( PdfUnits.MmToPoints(height));
        }

        public static PdfPageSize A5 => new PdfPageSize(148d, 210d);
        public static PdfPageSize A4 => new PdfPageSize(210d, 297d); //(595, 842);
        public static PdfPageSize A3 => new PdfPageSize(297d, 420d); //(842, 1191);
        public static PdfPageSize B5 => new PdfPageSize(182d, 257d);
        public static PdfPageSize B4 => new PdfPageSize(257d, 364d);
        public static PdfPageSize Letter => new PdfPageSize(215.9d, 279.4d); //(612, 792);
        public static PdfPageSize Legal => new PdfPageSize(215.9d, 355.6d); //(612, 1008);
        public static PdfPageSize Statement => new PdfPageSize(139.7d, 215.9d);
        public static PdfPageSize Executive => new PdfPageSize(184.2d, 266.7d);
    }
}