using OfficeOpenXml.PDF.PdfGraphics;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfPatternDarkUp : PdfPatternFill
    {
        public PdfPatternDarkUp(PdfColor foreground, PdfColor background) : base(foreground, background) { }

        public override string CreatePatternResource()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"{Foreground.ToStrokeCommand()}\n" +
                            $"1.4 w\n" +
                            $"11.3137 0 m\n" +
                            $"0 11.3137 l\n" +
                            $"S");
            return sb.ToString();
        }
    }
}
