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
using EPPlus.Export.Pdf.Helpers;
using System.Drawing;

namespace EPPlus.Export.Pdf.Layout
{
    internal class PdfCellGradientFillData
    {
        public ExcelFillGradientType GradientType;
        public Color Color1;
        public Color Color2;
        public Color Color3;
        public double Degree;
        public double Top;
        public double Bottom;
        public double Left;
        public double Right;
        public double[] matrix;
        public double[] coords;

        public override string ToString()
        {
            return GradientType.ToString() + Color1.ToHexString() + Color2.ToHexString() + Degree + Top + Bottom + Left + Right;
        }
    }
}
