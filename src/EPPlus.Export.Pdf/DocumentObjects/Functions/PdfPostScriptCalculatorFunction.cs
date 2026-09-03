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
    /// <summary>
    /// FunctionType 4 (PostScript calculator). Maps a 2-D point (u,v) in the unit domain to an
    /// RGB colour using the "box" (rectangular) gradient parameter Excel uses for path gradients:
    /// t = max(|u-fx|/dx, |v-fy|/dy). Unlike the Type 2/3 functions, a Type 4 function is a
    /// stream object, so it must be an indirect object and be referenced by the shading (N 0 R).
    /// </summary>
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
            // Focus point (fx,fy) and half-extents (dx,dy) in the shading's unit domain (v is "up").
            GetFocus(g, out double fx, out double fy, out double dx, out double dy);

            // Colours are normalised to 0..1 for DeviceRGB. Color1 = focus (t=0), Color2 = edge (t=1).
            double r0 = g.Color1.GetR(), g0 = g.Color1.GetG(), b0 = g.Color1.GetB();
            double r1 = g.Color2.GetR(), g1 = g.Color2.GetG(), b1 = g.Color2.GetB();

            var sb = new StringBuilder();
            sb.Append("{ ");
            // Stack in: u v (v on top). Compute t = max(|u-fx|/dx, |v-fy|/dy), then clamp to [0,1].
            sb.Append($"{fy.ToPdfString()} sub abs {dy.ToPdfString()} div ");   // |v-fy|/dy
            sb.Append("exch ");
            sb.Append($"{fx.ToPdfString()} sub abs {dx.ToPdfString()} div ");   // |u-fx|/dx
            sb.Append("2 copy lt { exch } if pop ");                            // -> max
            sb.Append("dup 1 gt { pop 1 } if ");                               // clamp high (abs keeps >= 0)

            if (!g.Color3.Equals(Color.Empty))
            {
                double rm = g.Color3.GetR(), gm = g.Color3.GetG(), bm = g.Color3.GetB();
                sb.Append("dup 0.5 le { 2 mul ");             // t in [0,0.5] -> s = t*2
                AppendRamp(sb, r0, g0, b0, rm, gm, bm);       // Color1 -> Color3
                sb.Append("} { 0.5 sub 2 mul ");              // t in (0.5,1] -> s = (t-0.5)*2
                AppendRamp(sb, rm, gm, bm, r1, g1, b1);       // Color3 -> Color2
                sb.Append("} ifelse ");
            }
            else
            {
                AppendRamp(sb, r0, g0, b0, r1, g1, b1);        // Color1 -> Color2
            }
            sb.Append("}");
            return sb.ToString();
        }

        // Given parameter s in [0,1] on the stack, leave R G B where each = c0 + s*(c1 - c0).
        private static void AppendRamp(StringBuilder sb,
            double r0, double g0, double b0, double r1, double g1, double b1)
        {
            sb.Append($"dup {r1.ToPdfString()} {r0.ToPdfString()} sub mul {r0.ToPdfString()} add exch ");
            sb.Append($"dup {g1.ToPdfString()} {g0.ToPdfString()} sub mul {g0.ToPdfString()} add exch ");
            sb.Append($"{b1.ToPdfString()} {b0.ToPdfString()} sub mul {b0.ToPdfString()} add ");
        }

        // The five Excel presets (four corners + centre). fillToRect insets are stored in
        // Left/Right/Top/Bottom; Top==0 means the focus is at the top edge (v == 1 in unit space).
        private static void GetFocus(PdfCellGradientFillData g, out double fx, out double fy, out double dx, out double dy)
        {
            if (g.Left == 0.5 && g.Right == 0.5 && g.Top == 0.5 && g.Bottom == 0.5)
            {
                fx = 0.5; fy = 0.5; dx = 0.5; dy = 0.5;   // from centre
            }
            else
            {
                fx = g.Left == 0 ? 0d : 1d;               // left inset 0 -> focus at left edge
                fy = g.Top == 0 ? 1d : 0d;                // top inset 0  -> focus at top edge (v up)
                dx = 1d; dy = 1d;                         // from a corner
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
                           $"/Length {bytes.Length} >>\nstream\n");
            bw.Write(bytes);
            WriteAscii(bw, "\nendstream");
        }
    }
}