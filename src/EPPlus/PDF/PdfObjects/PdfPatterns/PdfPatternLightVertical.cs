using OfficeOpenXml.PDF.PdfGraphics;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfPatternLightVertical : PdfPatternFill
    {
        public PdfPatternLightVertical(PdfColor foreground, PdfColor background) : base(foreground, background)
        {
        }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Background.ToFillCommand()}\n" +
                            $"0 0 1 0.5 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0 0 0.5 0.5 re\n" +
                            $"f");
            return sb.ToString();
        }
    }
}
