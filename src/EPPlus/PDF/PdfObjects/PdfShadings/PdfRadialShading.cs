using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.PDF.PdfObjects.PdfFunctions;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfShadings
{
    internal class PdfRadialShading : PdfShading
    {
        internal double[] Coords;
        internal double[] Domain;
        internal List<PdfFunction> Functions = new List<PdfFunction>();
        internal bool[] Extend = null;

        public PdfRadialShading(int objectNumber, PdfCellGradientFillData GradientFillData, int version = 0)
            : base(objectNumber, version)
        {
            ColorSpace = DeviceColorSpace.DeviceRGB;
            Coords = [0, 0, 1, 0];
            var func = new PdfExponentialInterpolationFunction(0);
            func.C0 = [GradientFillData.Color1.R, GradientFillData.Color1.G, GradientFillData.Color1.B];
            func.C1 = [GradientFillData.Color2.R, GradientFillData.Color2.G, GradientFillData.Color2.B];
            func.Domain = [0, 1];
            func.N = 1;
            Functions.Add(func);
            Extend = [true, true];
        }

        internal override string RenderDictionary()
        {
            var coordsStr = string.Join(" ", Coords.Select(w => w.ToPdfString()).ToArray());
            var functionsStr = string.Join("\n", Functions.Select(w => w.RenderDictionary()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Shading\n" +
                            $"   /ShadingType 3\n" +
                            $"   /ColorSpace /{ColorSpace.ToString()}\n" +
                            $"   /Coords [ {coordsStr} ]\n");
            sb.AppendFormat($"   /Function {functionsStr}");
            if (Domain != null)
            {
                var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
                sb.AppendFormat($"\n   /Domain [ {domainStr} ]");
            }
            if (Extend != null)
            {
                var extendStr = string.Join(" ", Extend.Select(w => w.ToString().ToLower()).ToArray());
                sb.AppendFormat($"\n   /Extend [ {extendStr} ]");
            }
            sb.Append(" >>");
            return sb.ToString();
        }
    }
}
