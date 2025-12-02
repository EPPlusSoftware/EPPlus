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
using System.Drawing;
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfObjects.PdfFunctions;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfShadings
{
    internal class PdfAxialShading : PdfShading
    {
        internal double[] Coords;
        internal double[] Domain;
        internal double[] Matrix;
        internal PdfFunction Function;
        internal bool[] Extend = null;

        public PdfAxialShading(int objectNumber, PdfCellGradientFillData GradientFillData, int version = 0)
            : base(objectNumber, version)
        {
            //Fun test later for diagonal gradients. Diagonal gradients are set to 45 degrees, but in excel they go from corner to corner. To replicate this we could create out 45 in a square that we then scale to fill the dell.
            ColorSpace = DeviceColorSpace.DeviceRGB;
            Coords = [0, 0, 1, 0];
            if (!GradientFillData.Color3.Equals(Color.Empty))
            {
                var func = new PdfStichingFunction(0);
                func.Domain = [0d, 1d];
                func.Bounds = [0.5d];
                func.Encode = [0d, 1d, 1d, 0d];
                var f1 = new PdfExponentialInterpolationFunction(0);
                f1.C0 = [GradientFillData.Color1.R, GradientFillData.Color1.G, GradientFillData.Color1.B];
                f1.C1 = [GradientFillData.Color3.R, GradientFillData.Color3.G, GradientFillData.Color3.B];
                f1.Domain = [0, 1];
                f1.N = 1;
                func.Functions.Add(f1);
                var f2 = new PdfExponentialInterpolationFunction(0);
                f2.C0 = [GradientFillData.Color2.R, GradientFillData.Color2.G, GradientFillData.Color2.B];
                f2.C1 = [GradientFillData.Color3.R, GradientFillData.Color3.G, GradientFillData.Color3.B];
                f2.Domain = [0, 1];
                f2.N = 1;
                func.Functions.Add(f2);
                Function = func;
            }
            else
            {
                var func = new PdfExponentialInterpolationFunction(0);
                func.C0 = [GradientFillData.Color1.R, GradientFillData.Color1.G, GradientFillData.Color1.B];
                func.C1 = [GradientFillData.Color2.R, GradientFillData.Color2.G, GradientFillData.Color2.B];
                func.Domain = [0, 1];
                func.N = 1;
                Function = func;
            }
            //Matrix = GradientFillData.matrix;
            Extend = [true, true];
        }

        internal override string RenderDictionary()
        {
            var coordsStr = string.Join(" ", Coords.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Shading\n" +
                            $"   /ShadingType 2\n" +
                            $"   /ColorSpace /{ColorSpace.ToString()}\n" +
                            $"   /Coords [ {coordsStr} ]\n");
            sb.AppendFormat($"   /Function {Function.RenderDictionary()}");
            if (Domain != null)
            {
                var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
                sb.AppendFormat($"\n   /Domain [ {domainStr} ]");
            }
            //if (Matrix != null)
            //{
            //    var matrixStr = string.Join(" ", Matrix.Select(w => w.ToPdfString()).ToArray());
            //    sb.AppendFormat($"\n   /matrix [ {matrixStr} ]");
            //}
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
