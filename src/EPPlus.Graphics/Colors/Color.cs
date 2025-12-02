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
using System.Globalization;

namespace EPPlus.Graphics.Colors
{
    internal class Color
    {
        public float R { get; set; }
        public float G { get; set; }
        public float B { get; set; }
        public float A { get; set; } = 1f;

        public Color() { }

        public Color(float r, float g, float b)
        {
            R = r;
            G = g;
            B = b;
        }

        public Color(float r, float g, float b, float a)
        {
            R = r;
            G = g;
            B = b;
            A = a;
        }

        public Color(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex == "0")
            {
                R = 0;
                G = 0;
                B = 0;
                A = 0;
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

        public bool Equals(Color other)
        {
            if (other is null) return false;
            return R == other.R && G == other.G && B == other.B && A == other.A;
        }

        public override bool Equals(object obj) => Equals(obj as Color);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 31 + R.GetHashCode();
                hash = hash * 31 + G.GetHashCode();
                hash = hash * 31 + B.GetHashCode();
                hash = hash * 31 + A.GetHashCode();
                return hash;
            }
        }

        public string ToHexString()
        {
            int r = (int)(R * 255);
            int g = (int)(G * 255);
            int b = (int)(B * 255);
            int a = (int)(A * 255);
            return $"#{r:X2}{g:X2}{b:X2}{a:X2}";
        }

        public static Color Red => new Color(1, 0, 0);
        public static Color Green => new Color(0, 1, 0);
        public static Color Blue => new Color(0, 0, 1);
        public static Color Black => new Color(0, 0, 0);
        public static Color White => new Color(1, 1, 1);
        public static Color Gray => new Color(0.5f, 0.5f, 0.5f);
        public static Color LightGray => new Color(0.75f, 0.75f, 0.75f);
        public static Color None => new Color(0, 0, 0, 0);
    }
}
