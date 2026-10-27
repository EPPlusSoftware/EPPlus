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
using System;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.Helpers
{
    internal class JpegCmykHelper
    {
        internal static bool TryDecode(byte[] jpeg, out int width, out int height, out byte[] gray)
        {
            width = 0; height = 0; gray = null;
            try { return Decode(jpeg, out width, out height, out gray); }
            catch { return false; }
        }

        private static readonly int[] ZZ = {
            0,1,8,16,9,2,3,10,17,24,32,25,18,11,4,5,12,19,26,33,40,48,41,34,27,20,13,6,7,14,21,28,
            35,42,49,56,57,50,43,36,29,22,15,23,30,37,44,51,58,59,52,45,38,31,39,46,53,60,61,54,47,55,62,63 };

        private static readonly double[,] CC = BuildIdctBasis();
        private static double[,] BuildIdctBasis()
        {
            var c = new double[8, 8];
            for (int u = 0; u < 8; u++)
            {
                double cu = u == 0 ? Math.Sqrt(1.0 / 8.0) : Math.Sqrt(2.0 / 8.0);
                for (int x = 0; x < 8; x++) c[u, x] = cu * Math.Cos((2 * x + 1) * u * Math.PI / 16.0);
            }
            return c;
        }

        private sealed class HDec
        {
            public int[] MinCode = new int[17], MaxCode = new int[17], ValPtr = new int[17];
            public byte[] Vals;
            public static HDec Build(byte[] bits1to16, byte[] vals)
            {
                var h = new HDec { Vals = vals };
                var sizes = new List<int>();
                for (int l = 1; l <= 16; l++) for (int n = 0; n < bits1to16[l]; n++) sizes.Add(l);
                var codes = new int[sizes.Count];
                int code = 0, k = 0;
                if (sizes.Count > 0)
                {
                    int si = sizes[0];
                    while (k < sizes.Count)
                    {
                        while (k < sizes.Count && sizes[k] == si) { codes[k] = code; code++; k++; }
                        code <<= 1; si++;
                    }
                }
                k = 0;
                for (int l = 1; l <= 16; l++)
                {
                    if (bits1to16[l] == 0) h.MaxCode[l] = -1;
                    else { h.ValPtr[l] = k; h.MinCode[l] = codes[k]; k += bits1to16[l]; h.MaxCode[l] = codes[k - 1]; }
                }
                return h;
            }
        }

        private sealed class BitR
        {
            private readonly byte[] d; private int pos; private readonly int end; private int buf, cnt;
            public BitR(byte[] data, int start, int stop) { d = data; pos = start; end = stop; }
            public int Bit()
            {
                if (cnt == 0)
                {
                    int b = pos < end ? (d[pos] & 0xFF) : 0; pos++;
                    if (b == 0xFF && pos < d.Length && (d[pos] & 0xFF) == 0x00) pos++;
                    buf = b; cnt = 8;
                }
                cnt--; return (buf >> cnt) & 1;
            }
            public int Bits(int n) { int v = 0; while (n-- > 0) v = (v << 1) | Bit(); return v; }
            public void Align() { cnt = 0; }
        }

        private sealed class Comp { public int Id, H, V, Q, Pred; public double[] Plane; public int Pw, Ph; }

        private static int Extend(int v, int s) => v >= (1 << (s - 1)) ? v : v - (1 << s) + 1;

        private static int Decode(BitR br, HDec t)
        {
            int code = br.Bit(), l = 1;
            while (l <= 16 && (t.MaxCode[l] < 0 || code > t.MaxCode[l])) { code = (code << 1) | br.Bit(); l++; }
            if (l > 16) return 0;
            return t.Vals[t.ValPtr[l] + code - t.MinCode[l]] & 0xFF;
        }

        private static bool Decode(byte[] d, out int width, out int height, out byte[] gray)
        {
            width = 0; height = 0; gray = null;
            if (d == null || d.Length < 4 || (d[0] & 0xFF) != 0xFF || (d[1] & 0xFF) != 0xD8) return false;

            var qt = new int[4][];
            var dc = new HDec[4]; var ac = new HDec[4];
            var comps = new List<Comp>();
            int W = 0, H = 0, dri = 0, adobeTransform = -1; bool adobeSeen = false;
            int scanStart = -1, scanEnd = -1;
            var sel = new List<(int sof, int td, int ta)>();

            int i = 2;
            while (i < d.Length - 1)
            {
                if ((d[i] & 0xFF) != 0xFF) { i++; continue; }
                int m = d[i + 1] & 0xFF;
                if (m == 0xFF) { i++; continue; }
                if (m == 0xD8) { i += 2; continue; }
                if (m == 0xD9) break;
                int seg = ((d[i + 2] & 0xFF) << 8) | (d[i + 3] & 0xFF);
                int p = i + 4, segEnd = i + 2 + seg;

                if (m == 0xC0 || m == 0xC1)                       // baseline / ext-sequential only
                {
                    H = ((d[p + 1] & 0xFF) << 8) | (d[p + 2] & 0xFF);
                    W = ((d[p + 3] & 0xFF) << 8) | (d[p + 4] & 0xFF);
                    int nc = d[p + 5] & 0xFF;
                    if (nc != 4) return false;                    // only CMYK here
                    for (int c = 0; c < nc; c++)
                    {
                        int q = p + 6 + c * 3;
                        comps.Add(new Comp { Id = d[q] & 0xFF, H = (d[q + 1] >> 4) & 0xF, V = d[q + 1] & 0xF, Q = d[q + 2] & 0xFF });
                    }
                }
                else if (m == 0xC2 || m == 0xC3 || (m >= 0xC5 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF))
                    return false;                                 // progressive / arithmetic: not handled
                else if (m == 0xDB)
                {
                    int q = p;
                    while (q < segEnd)
                    {
                        int pq = (d[q] >> 4) & 0xF, tq = d[q] & 0xF; q++;
                        var tbl = new int[64];
                        if (pq == 0) { for (int n = 0; n < 64; n++) tbl[n] = d[q++] & 0xFF; }
                        else { for (int n = 0; n < 64; n++) { tbl[n] = ((d[q] & 0xFF) << 8) | (d[q + 1] & 0xFF); q += 2; } }
                        qt[tq] = tbl;
                    }
                }
                else if (m == 0xC4)
                {
                    int q = p;
                    while (q < segEnd)
                    {
                        int tc = (d[q] >> 4) & 0xF, th = d[q] & 0xF; q++;
                        var bits = new byte[17]; int tot = 0;
                        for (int l = 1; l <= 16; l++) { bits[l] = d[q++]; tot += bits[l]; }
                        var vals = new byte[tot];
                        for (int v = 0; v < tot; v++) vals[v] = d[q++];
                        if (tc == 0) dc[th] = HDec.Build(bits, vals); else ac[th] = HDec.Build(bits, vals);
                    }
                }
                else if (m == 0xDD) { dri = ((d[p] & 0xFF) << 8) | (d[p + 1] & 0xFF); if (dri > 0) return false; } // restart: bail (rare for CMYK)
                else if (m == 0xEE && seg >= 14 &&
                         d[p] == (byte)'A' && d[p + 1] == (byte)'d' && d[p + 2] == (byte)'o' && d[p + 3] == (byte)'b' && d[p + 4] == (byte)'e')
                {
                    adobeSeen = true; adobeTransform = d[p + 11] & 0xFF;
                }
                else if (m == 0xDA)
                {
                    int ns = d[p] & 0xFF;
                    for (int s = 0; s < ns; s++)
                    {
                        int cid = d[p + 1 + s * 2] & 0xFF, tdta = d[p + 2 + s * 2] & 0xFF;
                        int sof = 0; for (int c = 0; c < comps.Count; c++) if (comps[c].Id == cid) sof = c;
                        sel.Add((sof, (tdta >> 4) & 0xF, tdta & 0xF));
                    }
                    scanStart = p + 1 + ns * 2 + 3;
                    int j = scanStart;
                    while (j < d.Length - 1)
                    {
                        if ((d[j] & 0xFF) == 0xFF)
                        {
                            int mm = d[j + 1] & 0xFF;
                            if (mm != 0x00 && !(mm >= 0xD0 && mm <= 0xD7)) break;
                        }
                        j++;
                    }
                    scanEnd = j; break;
                }
                i = segEnd;
            }

            if (W == 0 || H == 0 || comps.Count != 4 || scanStart < 0) return false;
            if (adobeTransform == 2) return false;                // YCCK -> handled by the transcoder

            int Hmax = 1, Vmax = 1;
            foreach (var c in comps) { if (c.H > Hmax) Hmax = c.H; if (c.V > Vmax) Vmax = c.V; }
            int mcuW = (W + 8 * Hmax - 1) / (8 * Hmax);
            int mcuH = (H + 8 * Vmax - 1) / (8 * Vmax);
            foreach (var c in comps)
            {
                c.Pw = mcuW * c.H * 8; c.Ph = mcuH * c.V * 8;
                c.Plane = new double[c.Pw * c.Ph];
            }

            var br = new BitR(d, scanStart, scanEnd);
            var coef = new double[64]; var nat = new double[64]; var spa = new double[64];
            int mcuCount = 0;
            for (int my = 0; my < mcuH; my++)
                for (int mx = 0; mx < mcuW; mx++)
                {
                    if (dri > 0 && mcuCount > 0 && mcuCount % dri == 0)
                    {
                        br.Align();
                        // skip RSTn is implicit: BitR stops at end; markers inside handled by re-seek is omitted
                        foreach (var c in comps) c.Pred = 0;
                    }
                    foreach (var (sof, td, ta) in sel)
                    {
                        var c = comps[sof];
                        for (int v = 0; v < c.V; v++)
                            for (int h = 0; h < c.H; h++)
                            {
                                Array.Clear(coef, 0, 64);
                                int s = Decode(br, dc[td]);
                                int diff = s != 0 ? Extend(br.Bits(s), s) : 0;
                                c.Pred += diff; coef[0] = c.Pred;
                                int k = 1;
                                while (k < 64)
                                {
                                    int rs = Decode(br, ac[ta]), r = rs >> 4, sz = rs & 0xF;
                                    if (sz == 0) { if (r == 15) { k += 16; continue; } break; }
                                    k += r; if (k > 63) break;
                                    coef[k] = Extend(br.Bits(sz), sz); k++;
                                }
                                var q = qt[c.Q];
                                Array.Clear(nat, 0, 64);
                                for (int zz = 0; zz < 64; zz++) nat[ZZ[zz]] = coef[zz] * q[zz];
                                Idct8x8(nat, spa);
                                int by = (my * c.V + v) * 8, bx = (mx * c.H + h) * 8;
                                for (int yy = 0; yy < 8; yy++)
                                {
                                    int row = (by + yy) * c.Pw + bx;
                                    for (int xx = 0; xx < 8; xx++) c.Plane[row + xx] = spa[yy * 8 + xx] + 128.0;
                                }
                            }
                    }
                    mcuCount++;
                }

            // assemble gray
            gray = new byte[W * H];
            var Cp = comps[0]; var Mp = comps[1]; var Yp = comps[2]; var Kp = comps[3];
            for (int y = 0; y < H; y++)
                for (int x = 0; x < W; x++)
                {
                    double cv = Sample(Cp, x, y, Hmax, Vmax);
                    double mv = Sample(Mp, x, y, Hmax, Vmax);
                    double yv = Sample(Yp, x, y, Hmax, Vmax);
                    double kv = Sample(Kp, x, y, Hmax, Vmax);
                    if (adobeSeen) { cv = 255 - cv; mv = 255 - mv; yv = 255 - yv; kv = 255 - kv; }
                    double r = (255 - cv) * (255 - kv) / 255.0;
                    double g = (255 - mv) * (255 - kv) / 255.0;
                    double b = (255 - yv) * (255 - kv) / 255.0;
                    int lum = (int)(0.2126 * r + 0.7152 * g + 0.0722 * b + 0.5);
                    gray[y * W + x] = (byte)(lum < 0 ? 0 : lum > 255 ? 255 : lum);
                }
            width = W; height = H;
            return true;
        }

        private static double Sample(Comp c, int x, int y, int Hmax, int Vmax)
        {
            int sx = x * c.H / Hmax, sy = y * c.V / Vmax;   // nearest-neighbour upsample
            if (sx >= c.Pw) sx = c.Pw - 1;
            if (sy >= c.Ph) sy = c.Ph - 1;
            double v = c.Plane[sy * c.Pw + sx];
            return v < 0 ? 0 : v > 255 ? 255 : v;
        }

        private static void Idct8x8(double[] f, double[] outp)
        {
            // spatial = C^T * F * C,  with F stored row-major natural order
            var tmp = new double[64];
            for (int x = 0; x < 8; x++)                       // tmp = C^T * F
                for (int j = 0; j < 8; j++)
                {
                    double sum = 0;
                    for (int u = 0; u < 8; u++) sum += CC[u, x] * f[u * 8 + j];
                    tmp[x * 8 + j] = sum;
                }
            for (int xr = 0; xr < 8; xr++)                    // out = tmp * C
                for (int xc = 0; xc < 8; xc++)
                {
                    double sum = 0;
                    for (int v = 0; v < 8; v++) sum += tmp[xr * 8 + v] * CC[v, xc];
                    outp[xr * 8 + xc] = sum;
                }
        }
    }
}
