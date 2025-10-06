using OfficeOpenXml.PDF.PdfGraphics;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfPatternDarkDown : PdfPatternFill
    {
        public PdfPatternDarkDown(PdfColor foreground, PdfColor background) : base(foreground, background) { }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Background.ToFillCommand()}\n" +
                            $"0 0 2 2 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0 0 0.5 0.5 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"1.5 0 0.5 0.5 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"1 0.5 1 0.5 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0.5 1 1 0.5 re\n" +
                            $"f\n" +
                            $"{Foreground.ToFillCommand()}\n" +
                            $"0 1.5 1 0.5 re\n" +
                            $"f");
            return sb.ToString();
        }
    }
}
