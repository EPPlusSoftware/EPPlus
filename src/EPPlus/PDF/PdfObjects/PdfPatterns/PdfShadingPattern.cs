using OfficeOpenXml.PDF.Pdfhelpers;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal  class PdfShadingPattern : PdfObject
    {
        internal int shadingObjectNumber;
        internal double[] Matrix = [51.8125, 0, 0, 15.0625, 50.4879429133858, 772.889517716535];
        //internal ExtGState dictionary //implement later

        public PdfShadingPattern(int objectNumber, int shadingObjectNumber, int version = 0)
            : base(objectNumber, version)
        {
            this.shadingObjectNumber = shadingObjectNumber;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Pattern\n" +
                            $"   /PatternType 2\n" +
                            $"   /Shading {shadingObjectNumber} 0 R");
            if (Matrix != null)
            {
                var matrixStr = string.Join(" ", Matrix.Select(w => w.ToPdfString()).ToArray());
                sb.AppendFormat($"\n   /Matrix [ {matrixStr} ]");
            }
            sb.Append(" >>");
            return sb.ToString();
        }
    }
}
