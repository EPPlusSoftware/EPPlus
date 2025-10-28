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
using EPPlus.Export.Pdf.PdfGraphics;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfPatterns
{
    internal class PdfPatternDarkDown : PdfPatternFill
    {
        public PdfPatternDarkDown(PdfColor foreground, PdfColor background) : base(foreground, background) { }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Foreground.ToStrokeCommand()}\n" +
                            $"1.4 w\n" +
                            $"0 11.3137 m\n" +
                            $"11.3137 0 l\n" +
                            $"S");
            return sb.ToString();
        }
    }
}
