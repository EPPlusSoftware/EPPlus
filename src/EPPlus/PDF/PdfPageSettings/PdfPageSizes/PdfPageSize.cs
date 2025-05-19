using OfficeOpenXml.PDF.PdfPageSettings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes
{
    public class PdfPageSize
    {
        public float Width { get; }
        public float Height { get; }

        public int WidthPoints { get; }
        public int HeightPoints { get; }


        public PdfPageSize(float width, float height)
        {
            Width = width;
            Height = height;
            WidthPoints = PdfUnits.MmToPointsRounded(width);
            HeightPoints = PdfUnits.MmToPointsRounded(height);
        }

        public static PdfPageSize A5 => new PdfPageSize(148, 210);
        public static PdfPageSize A4 => new PdfPageSize(210, 297); //(595, 842);
        public static PdfPageSize A3 => new PdfPageSize(297, 420); //(842, 1191);
        public static PdfPageSize B5 => new PdfPageSize(182, 257);
        public static PdfPageSize B4 => new PdfPageSize(257, 364);
        public static PdfPageSize Letter => new PdfPageSize(215.9f, 279.4f); //(612, 792);
        public static PdfPageSize Legal => new PdfPageSize(215.9f, 355.6f); //(612, 1008);
        public static PdfPageSize Statement => new PdfPageSize(139.7f, 215.9f);
        public static PdfPageSize Executive => new PdfPageSize(184.2f, 266.7f);
    }
}