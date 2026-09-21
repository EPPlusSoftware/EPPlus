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
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace EPPlus.Export.Pdf.DocumentObjects.Shadings
{
    internal class PdfMeshShading : PdfShading
    {
        private const int BitsPerFlag = 8;
        private const int BitsPerCoordinate = 16;
        private const int BitsPerComponent = 8;

        private readonly byte[] _mesh;

        public PdfMeshShading(int objectNumber, PdfCellGradientFillData gradientFillData, int version = 0)
            : base(objectNumber, version)
        {
            ColorSpace = DeviceColorSpace.DeviceRGB;
            _mesh = BuildMesh(gradientFillData);
        }

        // Focus (fx,fy) and half-extents (dx,dy) in unit space for the 5 Excel
        // presets; anything else falls back to centre. v (y) is "up" in PDF space,
        // so Top==0 => focus at the top (fy = 1).
        private static void GetFocus(PdfCellGradientFillData d,
            out double fx, out double fy, out double dx, out double dy)
        {
            bool center = d.Left == 0.5 && d.Right == 0.5 && d.Top == 0.5 && d.Bottom == 0.5;
            if (center) { fx = 0.5; fy = 0.5; dx = 0.5; dy = 0.5; }
            else
            {
                fx = d.Left == 0 ? 0.0 : 1.0;
                fy = d.Top == 0 ? 1.0 : 0.0;
                dx = 1.0; dy = 1.0;
            }
        }

        private static double Clamp01(double v) => v < 0 ? 0 : (v > 1 ? 1 : v);

        // Iso-rectangle corners at parameter t, ordered BL, BR, TR, TL.
        private static double[][] IsoRect(double fx, double fy, double dx, double dy, double t)
        {
            double x0 = Clamp01(fx - t * dx), x1 = Clamp01(fx + t * dx);
            double y0 = Clamp01(fy - t * dy), y1 = Clamp01(fy + t * dy);
            return new[] { new[] { x0, y0 }, new[] { x1, y0 }, new[] { x1, y1 }, new[] { x0, y1 } };
        }

        private byte[] BuildMesh(PdfCellGradientFillData d)
        {
            GetFocus(d, out var fx, out var fy, out var dx, out var dy);

            // Colour stops: t = 0 at the focus, t = 1 at the edge.
            bool has3 = !d.Color3.Equals(Color.Empty);
            var stops = has3
                ? new (double t, Color c)[] { (0.0, d.Color1), (0.5, d.Color3), (1.0, d.Color2) }
                : new (double t, Color c)[] { (0.0, d.Color1), (1.0, d.Color2) };

            var buf = new List<byte>();
            for (int i = 0; i < stops.Length - 1; i++)
            {
                var (ta, ca) = stops[i];
                var (tb, cb) = stops[i + 1];
                var inner = IsoRect(fx, fy, dx, dy, ta);
                var outer = IsoRect(fx, fy, dx, dy, tb);

                // Two triangles per side of the frame; degenerate ones are skipped
                // (inner ring is a point, or a side coincides with the cell edge).
                for (int k = 0; k < 4; k++)
                {
                    int k1 = (k + 1) & 3;
                    AddTriangle(buf, inner[k], ca, inner[k1], ca, outer[k1], cb);
                    AddTriangle(buf, inner[k], ca, outer[k1], cb, outer[k], cb);
                }
            }
            return buf.ToArray();
        }

        private static void AddTriangle(List<byte> buf,
            double[] p0, Color c0, double[] p1, Color c1, double[] p2, Color c2)
        {
            double area = Math.Abs((p1[0] - p0[0]) * (p2[1] - p0[1])
                                 - (p2[0] - p0[0]) * (p1[1] - p0[1]));
            if (area < 1e-9) return; // drop zero-area triangles

            AddVertex(buf, p0, c0);
            AddVertex(buf, p1, c1);
            AddVertex(buf, p2, c2);
        }

        // Free-form vertex: flag(1) x(2) y(2) r(1) g(1) b(1). All flags = 0, so
        // every three consecutive vertices form an independent triangle.
        private static void AddVertex(List<byte> buf, double[] p, Color c)
        {
            buf.Add(0);
            AddUInt16(buf, p[0]);
            AddUInt16(buf, p[1]);
            buf.Add(ToByte(c.GetR())); // GetR/G/B are normalised 0..1 (DeviceRGB)
            buf.Add(ToByte(c.GetG()));
            buf.Add(ToByte(c.GetB()));
        }

        private static void AddUInt16(List<byte> buf, double unit)
        {
            int v = (int)Math.Round(Clamp01(unit) * 65535.0);
            buf.Add((byte)((v >> 8) & 0xFF)); // big-endian
            buf.Add((byte)(v & 0xFF));
        }

        private static byte ToByte(double c01) => (byte)Math.Round(Clamp01(c01) * 255.0);

        private string DictHeader() =>
            "<< /Type /Shading\n" +
            "   /ShadingType 4\n" +
            "   /ColorSpace /" + ColorSpace + "\n" +
            "   /BitsPerCoordinate " + BitsPerCoordinate + "\n" +
            "   /BitsPerComponent " + BitsPerComponent + "\n" +
            "   /BitsPerFlag " + BitsPerFlag + "\n" +
            "   /Decode [0 1 0 1 0 1 0 1 0 1]\n" +
            "   /Length " + _mesh.Length + " >>\n";

        internal override string RenderDictionary()
            => DictHeader() + "stream\n|BINARY DATA|\nendstream";

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, DictHeader() + "stream\n");
            bw.Write(_mesh);
            WriteAscii(bw, "\nendstream");
        }
    }
}