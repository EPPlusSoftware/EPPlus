/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Patterns
{
    internal class PdfTilingPattern : PdfPattern
    {
        internal PdfPatternFill fill;
        internal double[] BBox = [0d, 0d, 0d, 0d];
        internal double XStep = 0d;
        internal double YStep = 0d;
        internal double[] Matrix;

         public PdfTilingPattern(int objectNumber, PdfPatternFill patternFill, double[] BBox, double XStep, double YStep, int version = 0) : base(objectNumber, version)
        {
            fill = patternFill;
            this.BBox = BBox;
            this.XStep = XStep;
            this.YStep = YStep;
        }

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

        internal override void RenderDictionary(BinaryWriter bw)
        {
            //var bboxStr = string.Join(" ", BBox.Select(x => x.ToPdfString()).ToArray());
            //var sb = new StringBuilder();
            //sb.AppendFormat($"<< /Type /Pattern\n" +
            //                $"   /PatternType 1\n" +
            //                $"   /PaintType 1\n" +
            //                $"   /TilingType 1\n" +
            //                $"   /BBox [ {bboxStr} ]\n" +
            //                $"   /XStep {XStep.ToPdfString()}\n" +
            //                $"   /YStep {YStep.ToPdfString()}\n" +
            //                $"   /Resources << >>");
            //if (Matrix != null)
            //{
            //    var matrixStr = string.Join(" ", Matrix.Select(w => w.ToPdfStringF4()).ToArray());
            //    sb.AppendFormat($"\n   /Matrix [ {matrixStr} ]");
            //}
            //if (fill != null)
            //{
            //    var streamContent = fill.CreatePatternResource();
            //    var bytes = Encoding.ASCII.GetBytes(streamContent);
            //    sb.AppendFormat($"\n   /Length {bytes.Length}");
            //    sb.Append(" >>");
            //    sb.AppendFormat($"\nstream\n{streamContent}\nendstream");
            //}
            //else
            //{
            //    sb.Append(" >>");
            //}
            //WriteAscii(bw, sb.ToString());

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
                var body = PdfFlate.Compress(Encoding.ASCII.GetBytes(streamContent));
                sb.AppendFormat($"\n   /Filter /FlateDecode /Length {body.Length}");
                sb.Append(" >>");
                WriteAscii(bw, sb.ToString());
                WriteAscii(bw, "\nstream\n");
                bw.Write(body);
                WriteAscii(bw, "\nendstream");
            }
            else
            {
                sb.Append(" >>");
                WriteAscii(bw, sb.ToString());
            }
        }
    }
}
