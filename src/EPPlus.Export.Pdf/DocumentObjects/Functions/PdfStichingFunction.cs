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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Functions
{
    internal class PdfStichingFunction : PdfFunction
    {
        internal List<PdfFunction> Functions = new List<PdfFunction>();
        internal double[] Bounds;
        internal double[] Encode;

        public PdfStichingFunction(int objectNumber, int version = 0) : base(objectNumber, version) { }

        internal override string RenderDictionary()
        {
            var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
            var functionsStr = string.Join("\n", Functions.Select(w => w.RenderDictionary()).ToArray());
            var boundsStr = string.Join(" ", Bounds.Select(w => w.ToPdfString()).ToArray());
            var encodeStr = string.Join(" ", Encode.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /FunctionType 3\n" +
                            $"   /Domain [ {domainStr} ]\n");
            sb.AppendFormat($"   /Functions [ {functionsStr} ]\n");
            sb.AppendFormat($"   /Bounds [ {boundsStr} ]\n" +
                            $"   /Encode [ {encodeStr} ] >>");
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
            var functionsStr = string.Join("\n", Functions.Select(w => w.RenderDictionary()).ToArray());
            var boundsStr = string.Join(" ", Bounds.Select(w => w.ToPdfString()).ToArray());
            var encodeStr = string.Join(" ", Encode.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /FunctionType 3\n" +
                            $"   /Domain [ {domainStr} ]\n");
            sb.AppendFormat($"   /Functions [ {functionsStr} ]\n");
            sb.AppendFormat($"   /Bounds [ {boundsStr} ]\n" +
                            $"   /Encode [ {encodeStr} ] >>");
            WriteAscii(bw, sb.ToString());
        }
    }
}
