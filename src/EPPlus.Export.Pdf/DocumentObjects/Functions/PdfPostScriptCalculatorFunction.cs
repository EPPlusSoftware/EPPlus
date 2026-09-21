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
using System.Drawing;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Functions
{
    internal class PdfPostScriptCalculatorFunction : PdfFunction
    {
        private readonly string _code;

        public PdfPostScriptCalculatorFunction(int objectNumber, PdfCellGradientFillData gradientFillData, int version = 0)
            : base(objectNumber, version)
        {
            _code = BuildBoxGradientCode(gradientFillData);
        }

        private static string BuildBoxGradientCode(PdfCellGradientFillData g)
        {
            GetFocus(g, out double fx, out double fy, out double dx, out double dy);
            double r0 = g.Color1.GetR(), g0 = g.Color1.GetG(), b0 = g.Color1.GetB();
            double r1 = g.Color2.GetR(), g1 = g.Color2.GetG(), b1 = g.Color2.GetB();
            var sb = new StringBuilder();
            sb.Append("{ ");
            sb.Append($"{fy.ToPdfString()} sub abs {dy.ToPdfString()} div ");
            sb.Append("exch ");
            sb.Append($"{fx.ToPdfString()} sub abs {dx.ToPdfString()} div ");
            sb.Append("2 copy lt { exch } if pop ");
            sb.Append("dup 1 gt { pop 1 } if ");
            if (!g.Color3.Equals(Color.Empty))
            {
                double rm = g.Color3.GetR(), gm = g.Color3.GetG(), bm = g.Color3.GetB();
                sb.Append("dup 0.5 le { 2 mul ");
                AppendRamp(sb, r0, g0, b0, rm, gm, bm);
                sb.Append("} { 0.5 sub 2 mul ");
                AppendRamp(sb, rm, gm, bm, r1, g1, b1);
                sb.Append("} ifelse ");
            }
            else
            {
                AppendRamp(sb, r0, g0, b0, r1, g1, b1);
            }
            sb.Append("}");
            return sb.ToString();
        }

        private static void AppendRamp(StringBuilder sb,
            double r0, double g0, double b0, double r1, double g1, double b1)
        {
            sb.Append($"dup {r1.ToPdfString()} {r0.ToPdfString()} sub mul {r0.ToPdfString()} add exch ");
            sb.Append($"dup {g1.ToPdfString()} {g0.ToPdfString()} sub mul {g0.ToPdfString()} add exch ");
            sb.Append($"{b1.ToPdfString()} {b0.ToPdfString()} sub mul {b0.ToPdfString()} add ");
        }

        private static void GetFocus(PdfCellGradientFillData g, out double fx, out double fy, out double dx, out double dy)
        {
            if (g.Left == 0.5 && g.Right == 0.5 && g.Top == 0.5 && g.Bottom == 0.5)
            {
                fx = 0.5; fy = 0.5; dx = 0.5; dy = 0.5;
            }
            else
            {
                fx = g.Left == 0 ? 0d : 1d;
                fy = g.Top == 0 ? 1d : 0d;
                dx = 1d; dy = 1d;
            }
        }

        internal override string RenderDictionary()
        {
            return "<< /FunctionType 4 /Domain [ 0 1 0 1 ] /Range [ 0 1 0 1 0 1 ] " +
                   $"/Length {Encoding.ASCII.GetByteCount(_code)} >>\nstream\n{_code}\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var bytes = Encoding.ASCII.GetBytes(_code);
            WriteAscii(bw, "<< /FunctionType 4 /Domain [ 0 1 0 1 ] /Range [ 0 1 0 1 0 1 ] " +
                           $"/Length {bytes.Length.ToPdfStringF0()} >>\nstream\n");
            bw.Write(bytes);
            WriteAscii(bw, "\nendstream");
        }
    }
}