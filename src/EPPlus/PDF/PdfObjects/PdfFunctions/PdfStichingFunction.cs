using OfficeOpenXml.PDF.Pdfhelpers;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfFunctions
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
    }
}
