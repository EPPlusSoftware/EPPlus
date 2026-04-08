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
using EPPlus.Export.Pdf.Pdfhelpers;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFunctions
{
    internal class PdfExponentialInterpolationFunction : PdfFunction
    {
        internal double[] C0;
        internal double[] C1;
        internal double N;

        public PdfExponentialInterpolationFunction(int objectNumber, int version = 0) : base(objectNumber, version) { }

        internal override string RenderDictionary()
        {
            var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
            var c0Str = string.Join(" ", C0.Select(w => w.ToPdfString()).ToArray());
            var c1Str = string.Join(" ", C1.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /FunctionType 2\n" +
                            $"   /Domain [ {domainStr} ]\n" +
                            $"   /C0 [ {c0Str} ]\n" +
                            $"   /C1 [ {c1Str} ]\n" +
                            $"   /N {N.ToPdfString()} >>");
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
            var c0Str = string.Join(" ", C0.Select(w => w.ToPdfString()).ToArray());
            var c1Str = string.Join(" ", C1.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /FunctionType 2\n" +
                            $"   /Domain [ {domainStr} ]\n" +
                            $"   /C0 [ {c0Str} ]\n" +
                            $"   /C1 [ {c1Str} ]\n" +
                            $"   /N {N.ToPdfString()} >>");
            WriteAscii(bw, sb.ToString());
        }
    }
}
