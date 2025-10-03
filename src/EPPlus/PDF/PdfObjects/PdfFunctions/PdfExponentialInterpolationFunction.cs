using OfficeOpenXml.PDF.Pdfhelpers;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfFunctions
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
    }
}
