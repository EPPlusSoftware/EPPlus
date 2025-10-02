using OfficeOpenXml.PDF.PdfGraphics;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfPatternLightTrellis : PdfPatternFill
    {
        public PdfPatternLightTrellis(PdfColor foreground, PdfColor background) : base(foreground, background)
        {
        }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Background.ToFillCommand()}\n" +
                            $"0 0 4 4 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0 0 3 1 re\n" +
                            $"1 1 1 1 re\n" +
                            $"3 1 1 1 re\n" +
                            $"0 2 1 1 re\n" +
                            $"2 2 2 1 re\n" +
                            $"1 3 1 1 re\n" +
                            $"3 3 1 1 re\n" +
                            $"f");
            return sb.ToString();
        }
    }
}
