using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal  class PdfShadingPattern : PdfObject
    {
        internal int shadingObjectNumber;
        internal double[] Matrix; //implement later
        //internal ExtGState dictionary //implement later

        public PdfShadingPattern(int objectNumber, int shadingObjectNumber, int version = 0)
            : base(objectNumber, version)
        {
            this.shadingObjectNumber = shadingObjectNumber;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type Pattern\n" +
                            $"   /PatternType 2\n" +
                            $"   /Shading {shadingObjectNumber} 0 R >>");
            return sb.ToString();
        }
    }
}
