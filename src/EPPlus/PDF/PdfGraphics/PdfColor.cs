using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfGraphics
{
    internal class PdfColor
    {
        public float R { get; set; }
        public float G { get; set; }
        public float B { get; set; }
        public float A { get; set; } = 1f;


        public PdfColor()
        { }

        public PdfColor(float r, float g, float b)
        {
            R = r;
            G = g;
            B = b;
        }
        public PdfColor(float r, float g, float b, float a)
        {
            R = r;
            G = g;
            B = b;
            A = a;
        }

        public PdfColor(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex == "0")
            {
                R = 1;
                G = 1;
                B = 1;
                A = 1;
                return;
            }

            hex = hex.Trim().TrimStart('#');

            if (hex.Length == 3)
            {
                R = Convert.ToByte(new string(hex[0], 2), 16) / 255f;
                G = Convert.ToByte(new string(hex[1], 2), 16) / 255f;
                B = Convert.ToByte(new string(hex[2], 2), 16) / 255f;
            }
            else if (hex.Length == 4)
            {
                A = Convert.ToByte(new string(hex[0], 2), 16) / 255f;
                R = Convert.ToByte(new string(hex[1], 2), 16) / 255f;
                G = Convert.ToByte(new string(hex[2], 2), 16) / 255f;
                B = Convert.ToByte(new string(hex[3], 2), 16) / 255f;
            }
            else if (hex.Length == 6)
            {
                R = Convert.ToByte(hex.Substring(0, 2), 16) / 255f;
                G = Convert.ToByte(hex.Substring(2, 2), 16) / 255f;
                B = Convert.ToByte(hex.Substring(4, 2), 16) / 255f;
            }
            else if (hex.Length == 8)
            {
                A = Convert.ToByte(hex.Substring(0, 2), 16) / 255f;
                R = Convert.ToByte(hex.Substring(2, 2), 16) / 255f;
                G = Convert.ToByte(hex.Substring(4, 2), 16) / 255f;
                B = Convert.ToByte(hex.Substring(6, 2), 16) / 255f;
            }
            else
            {
                throw new FormatException("Invalid hex color format.");
            }
        }

        public string ToStrokeCommand() => $"{R.ToString("F", CultureInfo.InvariantCulture)} {G.ToString("F", CultureInfo.InvariantCulture)} {B.ToString("F", CultureInfo.InvariantCulture)} RG";
        public string ToFillCommand() => $"{R.ToString("F", CultureInfo.InvariantCulture)} {G.ToString("F", CultureInfo.InvariantCulture)} {B.ToString("F", CultureInfo.InvariantCulture)} rg";

        public static PdfColor Red => new(1, 0, 0);
        public static PdfColor Green => new(0, 1, 0);
        public static PdfColor Blue => new(0, 0, 1);
        public static PdfColor Black => new(0, 0, 0);
        public static PdfColor White => new(1, 1, 1);
        public static PdfColor Gray => new(0.5f, 0.5f, 0.5f);
        public static PdfColor LightGray => new(0.75f, 0.75f, 0.75f);
        public static PdfColor None => new(0, 0, 0, 0);
    }
}
