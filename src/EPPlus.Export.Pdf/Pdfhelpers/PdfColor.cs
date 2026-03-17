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

namespace EPPlus.Export.Pdf.Pdfhelpers
{
    internal static class PdfColor
    {
        /// <summary>
        /// Returns the red component of the color as a nomralized value.
        /// </summary>
        /// <param name="c">The color to extract the component</param>
        /// <returns>Normalized value of the color component.</returns>
        public static float GetR(this Color c) => c.R / 255f;

        /// <summary>
        /// Returns the green component of the color as a nomralized value.
        /// </summary>
        /// <param name="c">The color to extract the component</param>
        /// <returns>Normalized value of the color component.</returns>
        public static float GetG(this Color c) => c.G / 255f;

        /// <summary>
        /// Returns the blue component of the color as a nomralized value.
        /// </summary>
        /// <param name="c">The color to extract the component</param>
        /// <returns>Normalized value of the color component.</returns>
        public static float GetB(this Color c) => c.B / 255f;

        /// <summary>
        /// Returns the alpha component of the color as a nomralized value.
        /// </summary>
        /// <param name="c">The color to extract the component</param>
        /// <returns>Normalized value of the color component.</returns>
        public static float GetA(this Color c) => c.A / 255f;

        /// <summary>
        /// Convert a hex string representing a color into a color.
        /// </summary>
        /// <param name="hex">String with corlor in hex code.
        /// Valid inputs that sets color to Color.Empty:
        /// "#0"
        /// "0"
        /// ""
        /// null
        /// </param>
        /// <returns>System.Drawing.Color object.</returns>
        /// <exception cref="FormatException">Throws exception of string is in invalid format.</exception>
        public static Color SetColorFromHex(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex == "0" || hex == "#0")
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
            return Color.FromArgb(A, R, G, B);
        }

        /// <summary>
        /// Returns a string for pdf command for stroke.
        /// </summary>
        /// <param name="c">The color to use for stroke.</param>
        /// <returns>The command string for a stroke color</returns>
        public static string ToStrokeCommand(this Color c) => $"{c.GetR().ToString("F", CultureInfo.InvariantCulture)} {c.GetG().ToString("F", CultureInfo.InvariantCulture)} {c.GetB().ToString("F", CultureInfo.InvariantCulture)} RG";

        /// <summary>
        /// Returns a string for pdf command for fill.
        /// </summary>
        /// <param name="c">The color to use for fill.</param>
        /// <returns>The command string for a fill color</returns>
        public static string ToFillCommand(this Color c) => $"{c.GetR().ToString("F", CultureInfo.InvariantCulture)} {c.GetG().ToString("F", CultureInfo.InvariantCulture)} {c.GetB().ToString("F", CultureInfo.InvariantCulture)} rg";

        /// <summary>
        /// Get the color represented in Hex as a string.
        /// </summary>
        /// <param name="c">The color which to return the hex value as a string.</param>
        /// <returns>String represeting the hex value of the color.</returns>
        public static string ToHexString(this Color c)
        {
            int r = (int)(c.R);
            int g = (int)(c.G);
            int b = (int)(c.B);
            int a = (int)(c.A);
            return $"#{r:X2}{g:X2}{b:X2}{a:X2}";
        }
    }
}
