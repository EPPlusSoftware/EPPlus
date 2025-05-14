using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfGraphics
{
    internal class PdfColor
    {
        public static PdfColor Black => new(0, 0, 0);
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

        public string ToStrokeCommand() => $"{R} {G} {B} RG";
        public string ToFillCommand() => $"{R} {G} {B} rg";
    }
}
