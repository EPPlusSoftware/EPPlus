using OfficeOpenXml.PDF.Pdfhelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal class PdfTilingPattern : PdfObject
    {
        internal PdfPatternFill fill;
        internal double[] Matrix;

        public PdfTilingPattern(int objectNumber, int version = 0) : base(objectNumber, version)
        {
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Pattern\n" +
                            $"   /PatternType 1\n" +
                            $"   /PaintType 1\n" +
                            $"   /TilingType 1" +
                            $"   /BBox [ {4} ]\n" +
                            $"   /XStep {1}\n" +
                            $"   /YStep {1}\n" +
                            $"   /Resources [ {0} ]");
            if (Matrix != null)
            {
                var matrixStr = string.Join(" ", Matrix.Select(w => w.ToPdfStringF4()).ToArray());
                sb.AppendFormat($"\n   /Matrix [ {matrixStr} ]");
            }
            var streamContent = fill.CreatePatternResource();
            var bytes = Encoding.ASCII.GetBytes(streamContent);
            sb.AppendFormat($"   /Length {bytes.Length}");
            sb.Append(" >>");
            sb.AppendFormat($"\nstream\n{streamContent}endstream");
            return sb.ToString();
        }
    }
}
