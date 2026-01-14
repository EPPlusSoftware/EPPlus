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
using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace EPPlus.Export.Pdf.Pdfhelpers
{
    internal static class PdfColor
    {
        public static float GetR(this Color c) => c.R / 255f;
        public static float GetG(this Color c) => c.G / 255f;
        public static float GetB(this Color c) => c.B / 255f;
        public static float GetA(this Color c) => c.A / 255f;

        public static Color SetColorFromHex(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex == "0")
            {
                return Color.Empty;

            }
            hex = hex.Trim().TrimStart('#');
            int R=0, G=0, B=0, A=0;
            if (hex.Length == 3)
            {
                R = Convert.ToByte(new string(hex[0], 2), 16);
                G = Convert.ToByte(new string(hex[1], 2), 16);
                B = Convert.ToByte(new string(hex[2], 2), 16);

            }
            else if (hex.Length == 4)
            {
                A = Convert.ToByte(new string(hex[0], 2), 16);
                R = Convert.ToByte(new string(hex[1], 2), 16);
                G = Convert.ToByte(new string(hex[2], 2), 16);
                B = Convert.ToByte(new string(hex[3], 2), 16);
            }
            else if (hex.Length == 6)
            {
                R = Convert.ToByte(hex.Substring(0, 2), 16);
                G = Convert.ToByte(hex.Substring(2, 2), 16);
                B = Convert.ToByte(hex.Substring(4, 2), 16);
            }
            else if (hex.Length == 8)
            {
                A = Convert.ToByte(hex.Substring(0, 2), 16);
                R = Convert.ToByte(hex.Substring(2, 2), 16);
                G = Convert.ToByte(hex.Substring(4, 2), 16);
                B = Convert.ToByte(hex.Substring(6, 2), 16);
            }
            else
            {
                throw new FormatException("Invalid hex color format.");
            }
            return Color.FromArgb(1, R, G, B);
        }

        public static string ToStrokeCommand(this Color c) => $"{c.GetR().ToString("F", CultureInfo.InvariantCulture)} {c.GetG().ToString("F", CultureInfo.InvariantCulture)} {c.GetB().ToString("F", CultureInfo.InvariantCulture)} RG";
        public static string ToFillCommand(this Color c) => $"{c.GetR().ToString("F", CultureInfo.InvariantCulture)} {c.GetG().ToString("F", CultureInfo.InvariantCulture)} {c.GetB().ToString("F", CultureInfo.InvariantCulture)} rg";

        //public bool Equals(Color other)
        //{
        //    if (other is null) return false;
        //    return R == other.R && G == other.G && B == other.B && A == other.A;
        //}

        //public override bool Equals(object obj) => Equals(obj as Color);

        //public override int GetHashCode()
        //{
        //    unchecked
        //    {
        //        int hash = 17;
        //        hash = hash * 31 + R.GetHashCode();
        //        hash = hash * 31 + G.GetHashCode();
        //        hash = hash * 31 + B.GetHashCode();
        //        hash = hash * 31 + A.GetHashCode();
        //        return hash;
        //    }
        //}

        public static string ToHexString(this Color c)
        {
            int r = (int)(c.R);
            int g = (int)(c.G);
            int b = (int)(c.B);
            int a = (int)(c.A);
            return $"#{r:X2}{g:X2}{b:X2}{a:X2}";
        }

        //public static Color Red => new Color(1, 0, 0);
        //public static Color Green => new Color(0, 1, 0);
        //public static Color Blue => new Color(0, 0, 1);
        //public static Color Black => new Color(0, 0, 0);
        //public static Color White => new Color(1, 1, 1);
        //public static Color Gray => new Color(0.5f, 0.5f, 0.5f);
        //public static Color LightGray => new Color(0.75f, 0.75f, 0.75f);
        //public static Color None => new Color(0, 0, 0, 0);
    }
}
