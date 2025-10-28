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
    public class PdfPageSize
    {
        public double Width { get; }
        public double Height { get; }
        public double WidthPu { get; }
        public double HeightPu { get; }

        public PdfPageSize(double width, double height)
        {
            Width = width;
            Height = height;
            WidthPu = System.Math.Round( PdfUnits.MmToPoints(width));
            HeightPu = System.Math.Round( PdfUnits.MmToPoints(height));
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