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
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfObjects.PdfFunctions;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfShadings
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
            func.C0 = [GradientFillData.Color1.GetR(), GradientFillData.Color1.GetG(), GradientFillData.Color1.GetB()];
            func.C1 = [GradientFillData.Color2.GetR(), GradientFillData.Color2.GetG(), GradientFillData.Color2.GetB()];
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

        internal override void RenderDictionary(BinaryWriter bw)
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
            WriteAscii(bw, sb.ToString());
        }
    }
}
