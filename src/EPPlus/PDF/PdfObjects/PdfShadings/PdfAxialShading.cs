using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.PDF.PdfObjects.PdfFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfShadings
{
    internal class PdfAxialShading : PdfShading
    {
        internal double[] Coords;
        internal double[] Domain = null;
        internal PdfFunction Function;
        internal bool[] Extend = null;

        public PdfAxialShading(int objectNumber, PdfCellGradientFillData GradientFillData, int version = 0)
            : base(objectNumber, version)
        {
            var cx = (GradientFillData.Left + GradientFillData.Right) / 2d;
            var cy = (GradientFillData.Top + GradientFillData.Bottom) / 2d;
            var halfLen = System.Math.Sqrt(System.Math.Pow(GradientFillData.Right - GradientFillData.Left, 2) + System.Math.Pow(GradientFillData.Top - GradientFillData.Bottom, 2)) / 2;
            var rad = GradientFillData.Degree * System.Math.PI / 180d;
            var dx = System.Math.Cos(rad) * halfLen;
            var dy = System.Math.Sin(rad) * halfLen;
            double x0 = cx - dx;
            double y0 = cy - dy;
            double x1 = cx + dx;
            double y1 = cy + dy;
            Coords = [x0, y0, x1, y1];
            var func = new PdfExponentialInterpolationFunction(0);
            func.C0 = [GradientFillData.Color0.R, GradientFillData.Color0.G, GradientFillData.Color0.B];
            func.C1 = [GradientFillData.Color1.R, GradientFillData.Color1.G, GradientFillData.Color1.B];
            func.Domain = [0, 1];
            func.N = 1;
            Function = func;
        }

        internal override string RenderDictionary()
        {
            var coordsStr = string.Join(" ", Coords.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /ShadingType 2\n" +
                            $"   /ColorSpace {ColorSpace.ToString()}\n" +
                            $"   /Coords [{coordsStr}]\n");
            sb.AppendFormat($"   /Function {Function.RenderDictionary()}");
            if (Domain != null)
            {
                var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
                sb.AppendFormat($"\n   /Domain [{domainStr}]");
            }
            if(Extend != null)
            {
                var extendStr = string.Join(" ", Extend.Select(w => w.ToString()).ToArray());
                sb.AppendFormat($"\n   /Extend [{extendStr}]");
            }
            sb.Append(">>");
            return sb.ToString();
        }
    }
}
