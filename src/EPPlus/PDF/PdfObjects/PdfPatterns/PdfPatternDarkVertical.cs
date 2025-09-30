using OfficeOpenXml.PDF.PdfGraphics;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfPatternDarkVertical : PdfPatternFill
    {
        public PdfPatternDarkVertical(PdfColor foreground, PdfColor background) : base(foreground, background) { }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Background.ToFillCommand()}\n" +
                            $"0 0 8 8 re\n" +
                            $"f\n" +
                            $"${Foreground.ToFillCommand()}\n" +
                            $"0 0 2 8 re\n" +
                            $"f");
            return sb.ToString();
        }
    }
}
