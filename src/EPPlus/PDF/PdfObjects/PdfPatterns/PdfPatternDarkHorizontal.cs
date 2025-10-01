using OfficeOpenXml.PDF.PdfGraphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfPatternDarkHorizontal : PdfPatternFill
    {
        public PdfPatternDarkHorizontal(PdfColor foreground, PdfColor background) : base(foreground, background) { }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Background.ToFillCommand()}\n" +
                            $"0 0 1 2 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0 0 1 1 re\n" +
                            $"f");
            return sb.ToString();
        }
    }
}
