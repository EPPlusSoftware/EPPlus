using OfficeOpenXml.PDF.Pdfhelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfTilingPattern : PdfPattern
    {
        internal PdfPatternFill fill;
        internal double[] BBox = [0d, 0d, 0d, 0d];
        internal double XStep = 0d;
        internal double YStep = 0d;
        internal double[] Matrix;


        public PdfTilingPattern(int objectNumber, int version = 0) : base(objectNumber, version) { }

        internal override string RenderDictionary()
        {
            var bboxStr = string.Join(" ", BBox.Select(x => x.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Pattern\n" +
                            $"   /PatternType 1\n" +
                            $"   /PaintType 1\n" +
                            $"   /TilingType 1\n" +
                            $"   /BBox [ {bboxStr} ]\n" +
                            $"   /XStep {XStep.ToPdfString()}\n" +
                            $"   /YStep {YStep.ToPdfString()}\n" +
                            $"   /Resources << >>");
            if (Matrix != null)
            {
                var matrixStr = string.Join(" ", Matrix.Select(w => w.ToPdfStringF4()).ToArray());
                sb.AppendFormat($"\n   /Matrix [ {matrixStr} ]");
            }
            if (fill != null)
            {
                var streamContent = fill.CreatePatternResource();
                var bytes = Encoding.ASCII.GetBytes(streamContent);
                sb.AppendFormat($"\n   /Length {bytes.Length}");
                sb.Append(" >>");
                sb.AppendFormat($"\nstream\n{streamContent}\nendstream");
            }
            else
            {
                sb.Append(" >>");
            }
            return sb.ToString();
        }
    }
}
