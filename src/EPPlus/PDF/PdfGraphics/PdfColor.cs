using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfGraphics
{
    public class PdfColor
    {
        public static PdfColor Black => new(0, 0, 0);
        public static PdfColor Gray => new(0.5f, 0.5f, 0.5f);
        public static PdfColor LightGray => new(0.75f, 0.75f, 0.75f);
        public static PdfColor White => new(1, 1, 1);
        public static PdfColor Red => new(1, 0, 0);
        public static PdfColor Green => new(0, 1, 0);
        public static PdfColor Blue => new(0, 0, 1);

        public float R { get; }
        public float G { get; }
        public float B { get; }

        public PdfColor(float r, float g, float b)
        {
            R = r;
            G = g;
            B = b;
        }

        public string ToStrokeCommand() => $"{R.ToString("F", CultureInfo.InvariantCulture)} {G.ToString("F", CultureInfo.InvariantCulture)} {B.ToString("F", CultureInfo.InvariantCulture)} RG";
        public string ToFillCommand() => $"{R.ToString("F", CultureInfo.InvariantCulture)} {G.ToString("F", CultureInfo.InvariantCulture)} {B.ToString("F", CultureInfo.InvariantCulture)} rg";
    }
}
