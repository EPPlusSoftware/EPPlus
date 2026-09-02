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
using EPPlus.Export.Pdf.Layout;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Shadings
{
    /// <summary>
    /// ShadingType 1 (function-based). The colour at each point comes from a 2-in / 3-out
    /// function of (u,v) evaluated over the unit Domain; the shading pattern's Matrix maps that
    /// unit square onto the cell. Used for Excel path (rectangular / "box") gradients, which have
    /// no native PDF shading. The Function is a stream object referenced indirectly.
    /// </summary>
    internal class PdfFunctionBasedShading : PdfShading
    {
        internal double[] Domain = [0d, 1d, 0d, 1d];
        internal int FunctionObjectNumber;

        public PdfFunctionBasedShading(int objectNumber, PdfCellGradientFillData gradientFillData, int version = 0)
            : base(objectNumber, version)
        {
            ColorSpace = DeviceColorSpace.DeviceRGB;
        }

        private string Build()
        {
            var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Shading\n" +
                            $"   /ShadingType 1\n" +
                            $"   /ColorSpace /{ColorSpace.ToString()}\n" +
                            $"   /Domain [ {domainStr} ]\n" +
                            $"   /Function {FunctionObjectNumber} 0 R >>");
            return sb.ToString();
        }

        internal override string RenderDictionary() => Build();

        internal override void RenderDictionary(BinaryWriter bw) => WriteAscii(bw, Build());
    }
}