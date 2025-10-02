using OfficeOpenXml.PDF.PdfGraphics;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfPatternLightGray : PdfPatternFill
    {
        public PdfPatternLightGray(PdfColor foreground, PdfColor background) : base(foreground, background)
        {
        }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Background.ToFillCommand()}\n" +
                            $"0 0 4 2 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0 1 1 1 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"2 0 1 1 re\n" +
                            $"f");
            return sb.ToString();
        }
    }
}
