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
using System.Drawing;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfPatterns
{
    internal class PdfPatternDarkTrellis : PdfPatternFill
    {
        public PdfPatternDarkTrellis(Color foreground, Color background) : base(foreground, background) { }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Background.ToFillCommand()}\n" +
                            $"0 0 4 4 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"2 1 2 1 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0 3 2 1 re\n" +
                            $"f");
            return sb.ToString();
        }
    }
}
