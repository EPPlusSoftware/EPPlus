using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Helpers
{
    internal class JpegProgressiveHelper
    {
        internal static byte[] ToGrayBaseline(byte[] jpeg)
        {
            try { return Transform(jpeg); }
            catch { return null; }
        }

        // ---- standard Annex K.3 Huffman tables (BITS[16] + VALUES) ----
        private static readonly byte[] DcLumBits = { 0, 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0 }; // [0] unused, 1..16
        private static readonly byte[] DcChrBits = { 0, 0, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0 };
        private static readonly byte[] DcVals = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 };
        private static readonly byte[] AcLumBits = { 0, 0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 0x7d };
        private static readonly byte[] AcChrBits = { 0, 0, 2, 1, 2, 4, 4, 3, 4, 7, 5, 4, 4, 0, 1, 2, 0x77 };
        private static readonly byte[] AcLumVals = {
            0x01,0x02,0x03,0x00,0x04,0x11,0x05,0x12,0x21,0x31,0x41,0x06,0x13,0x51,0x61,0x07,
            0x22,0x71,0x14,0x32,0x81,0x91,0xa1,0x08,0x23,0x42,0xb1,0xc1,0x15,0x52,0xd1,0xf0,
            0x24,0x33,0x62,0x72,0x82,0x09,0x0a,0x16,0x17,0x18,0x19,0x1a,0x25,0x26,0x27,0x28,
            0x29,0x2a,0x34,0x35,0x36,0x37,0x38,0x39,0x3a,0x43,0x44,0x45,0x46,0x47,0x48,0x49,
            0x4a,0x53,0x54,0x55,0x56,0x57,0x58,0x59,0x5a,0x63,0x64,0x65,0x66,0x67,0x68,0x69,
            0x6a,0x73,0x74,0x75,0x76,0x77,0x78,0x79,0x7a,0x83,0x84,0x85,0x86,0x87,0x88,0x89,
            0x8a,0x92,0x93,0x94,0x95,0x96,0x97,0x98,0x99,0x9a,0xa2,0xa3,0xa4,0xa5,0xa6,0xa7,
            0xa8,0xa9,0xaa,0xb2,0xb3,0xb4,0xb5,0xb6,0xb7,0xb8,0xb9,0xba,0xc2,0xc3,0xc4,0xc5,
            0xc6,0xc7,0xc8,0xc9,0xca,0xd2,0xd3,0xd4,0xd5,0xd6,0xd7,0xd8,0xd9,0xda,0xe1,0xe2,
            0xe3,0xe4,0xe5,0xe6,0xe7,0xe8,0xe9,0xea,0xf1,0xf2,0xf3,0xf4,0xf5,0xf6,0xf7,0xf8,
            0xf9,0xfa };
        private static readonly byte[] AcChrVals = {
            0x00,0x01,0x02,0x03,0x11,0x04,0x05,0x21,0x31,0x06,0x12,0x41,0x51,0x07,0x61,0x71,
            0x13,0x22,0x32,0x81,0x08,0x14,0x42,0x91,0xa1,0xb1,0xc1,0x09,0x23,0x33,0x52,0xf0,
            0x15,0x62,0x72,0xd1,0x0a,0x16,0x24,0x34,0xe1,0x25,0xf1,0x17,0x18,0x19,0x1a,0x26,
            0x27,0x28,0x29,0x2a,0x35,0x36,0x37,0x38,0x39,0x3a,0x43,0x44,0x45,0x46,0x47,0x48,
            0x49,0x4a,0x53,0x54,0x55,0x56,0x57,0x58,0x59,0x5a,0x63,0x64,0x65,0x66,0x67,0x68,
            0x69,0x6a,0x73,0x74,0x75,0x76,0x77,0x78,0x79,0x7a,0x82,0x83,0x84,0x85,0x86,0x87,
            0x88,0x89,0x8a,0x92,0x93,0x94,0x95,0x96,0x97,0x98,0x99,0x9a,0xa2,0xa3,0xa4,0xa5,
            0xa6,0xa7,0xa8,0xa9,0xaa,0xb2,0xb3,0xb4,0xb5,0xb6,0xb7,0xb8,0xb9,0xba,0xc2,0xc3,
            0xc4,0xc5,0xc6,0xc7,0xc8,0xc9,0xca,0xd2,0xd3,0xd4,0xd5,0xd6,0xd7,0xd8,0xd9,0xda,
            0xe2,0xe3,0xe4,0xe5,0xe6,0xe7,0xe8,0xe9,0xea,0xf2,0xf3,0xf4,0xf5,0xf6,0xf7,0xf8,
            0xf9,0xfa };

        private sealed class HDec
        {
            public int[] MinCode = new int[17];
            public int[] MaxCode = new int[17];
            public int[] ValPtr = new int[17];
            public byte[] Vals;
            public static HDec Build(byte[] bits1to16, byte[] vals) // bits indexed 1..16
            {
                var h = new HDec { Vals = vals };
                var sizes = new List<int>();
                for (int l = 1; l <= 16; l++) for (int i = 0; i < bits1to16[l]; i++) sizes.Add(l);
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
                    if (bits1to16[l] == 0) { h.MaxCode[l] = -1; }
                    else { h.ValPtr[l] = k; h.MinCode[l] = codes[k]; k += bits1to16[l]; h.MaxCode[l] = codes[k - 1]; }
                }
                return h;
            }
        }

        private sealed class HEnc
        {
            public int[] Code = new int[256];
            public int[] Len = new int[256];
            public static HEnc Build(byte[] bits1to16, byte[] vals)
            {
                var e = new HEnc();
                var sizes = new List<int>();
                for (int l = 1; l <= 16; l++) for (int i = 0; i < bits1to16[l]; i++) sizes.Add(l);
                int code = 0, k = 0;
                if (sizes.Count > 0)
                {
                    int si = sizes[0];
                    while (k < sizes.Count)
                    {
                        while (k < sizes.Count && sizes[k] == si) { e.Code[vals[k]] = code; e.Len[vals[k]] = si; code++; k++; }
                        code <<= 1; si++;
                    }
                }
                return e;
            }
        }

        private sealed class BitR
        {
            private readonly byte[] d; private int pos; private readonly int end;
            private int buf, cnt;
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
        }

        private sealed class Comp { public int Id, H, V, Q, Bpl, Bpc, WB, HB, Pred; public int[] Coef; }

        private static int Extend(int v, int s) => v >= (1 << (s - 1)) ? v : v - (1 << s) + 1;

        private static int Decode(BitR br, HDec t)
        {
            int code = br.Bit(), l = 1;
            while (l <= 16 && (t.MaxCode[l] < 0 || code > t.MaxCode[l])) { code = (code << 1) | br.Bit(); l++; }
            if (l > 16) return 0;
            return t.Vals[t.ValPtr[l] + code - t.MinCode[l]] & 0xFF;
        }

        private static byte[] Transform(byte[] d)
        {
            if (d == null || d.Length < 4 || (d[0] & 0xFF) != 0xFF || (d[1] & 0xFF) != 0xD8) return null;

            int W = 0, H = 0, Hmax = 1, Vmax = 1, mcuW = 0, mcuH = 0, dri = 0;
            var comps = new List<Comp>();
            var dqtSegs = new List<byte[]>();
            var dc = new HDec[4]; var ac = new HDec[4];
            int eobrun = 0;
            bool setup = false;

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

                if (m == 0xC2) // SOF2 progressive
                {
                    H = ((d[p + 1] & 0xFF) << 8) | (d[p + 2] & 0xFF);
                    W = ((d[p + 3] & 0xFF) << 8) | (d[p + 4] & 0xFF);
                    int nc = d[p + 5] & 0xFF;
                    if (nc != 3) return null;
                    for (int c = 0; c < nc; c++)
                    {
                        int q = p + 6 + c * 3;
                        comps.Add(new Comp { Id = d[q] & 0xFF, H = (d[q + 1] >> 4) & 0xF, V = d[q + 1] & 0xF, Q = d[q + 2] & 0xFF });
                    }
                    Hmax = 1; Vmax = 1;
                    foreach (var c in comps) { if (c.H > Hmax) Hmax = c.H; if (c.V > Vmax) Vmax = c.V; }
                    mcuW = (W + 8 * Hmax - 1) / (8 * Hmax);
                    mcuH = (H + 8 * Vmax - 1) / (8 * Vmax);
                    foreach (var c in comps)
                    {
                        c.Bpl = mcuW * c.H; c.Bpc = mcuH * c.V;
                        c.WB = (W * c.H + 8 * Hmax - 1) / (8 * Hmax);
                        c.HB = (H * c.V + 8 * Vmax - 1) / (8 * Vmax);
                        c.Coef = new int[c.Bpl * c.Bpc * 64];
                    }
                    setup = true;
                }
                else if (m == 0xC0 || m == 0xC1 || m == 0xC3 || (m >= 0xC5 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF))
                {
                    return null; // not progressive Huffman — not our job
                }
                else if (m == 0xDB) { var s = new byte[segEnd - i]; System.Array.Copy(d, i, s, 0, s.Length); dqtSegs.Add(s); }
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
                else if (m == 0xDD) { dri = ((d[p] & 0xFF) << 8) | (d[p + 1] & 0xFF); if (dri > 0) return null; } // progressive+restart: bail (rare)
                else if (m == 0xDA) // SOS — one scan
                {
                    if (!setup) return null;
                    int ns = d[p] & 0xFF;
                    var sel = new (int sof, int td, int ta)[ns];
                    for (int s = 0; s < ns; s++)
                    {
                        int cid = d[p + 1 + s * 2] & 0xFF, tdta = d[p + 2 + s * 2] & 0xFF;
                        int sof = 0; for (int c = 0; c < comps.Count; c++) if (comps[c].Id == cid) sof = c;
                        sel[s] = (sof, (tdta >> 4) & 0xF, tdta & 0xF);
                    }
                    int ss = d[p + 1 + ns * 2] & 0xFF, se = d[p + 2 + ns * 2] & 0xFF;
                    int ahal = d[p + 3 + ns * 2] & 0xFF, ah = ahal >> 4, al = ahal & 0xF;
                    int dataStart = p + 4 + ns * 2, j = dataStart;
                    while (j < d.Length - 1)
                    {
                        if ((d[j] & 0xFF) == 0xFF)
                        {
                            int mm = d[j + 1] & 0xFF;
                            if (mm != 0x00 && !(mm >= 0xD0 && mm <= 0xD7)) break;
                        }
                        j++;
                    }
                    DecodeScan(comps, sel, ss, se, ah, al, dc, ac, d, dataStart, j, mcuW, mcuH, ref eobrun);
                    i = j; continue;
                }
                i = segEnd;
            }

            if (!setup) return null;
            return EncodeBaseline(comps, dqtSegs, W, H, mcuW, mcuH);
        }

        private static void DecodeScan(List<Comp> comps, (int sof, int td, int ta)[] sel,
            int Ss, int Se, int Ah, int Al, HDec[] dc, HDec[] ac, byte[] d, int start, int end,
            int mcuW, int mcuH, ref int eobrun)
        {
            var br = new BitR(d, start, end);
            eobrun = 0;
            foreach (var c in comps) c.Pred = 0;

            if (Ss == 0) // DC scan (interleaved over scan components)
            {
                for (int my = 0; my < mcuH; my++)
                    for (int mx = 0; mx < mcuW; mx++)
                        foreach (var (sof, td, ta) in sel)
                        {
                            var c = comps[sof];
                            for (int v = 0; v < c.V; v++)
                                for (int h = 0; h < c.H; h++)
                                {
                                    int bRow = my * c.V + v, bCol = mx * c.H + h;
                                    int baseIdx = (bRow * c.Bpl + bCol) * 64;
                                    if (Ah == 0)
                                    {
                                        int s = Decode(br, dc[td]);
                                        int diff = s != 0 ? Extend(br.Bits(s), s) : 0;
                                        c.Pred += diff; c.Coef[baseIdx] = c.Pred << Al;
                                    }
                                    else if (br.Bit() != 0) c.Coef[baseIdx] |= (1 << Al);
                                }
                        }
                return;
            }

            // AC scan — single component, non-interleaved
            var comp = comps[sel[0].sof];
            var act = ac[sel[0].ta];
            int p1 = 1 << Al, m1 = (-1) << Al;

            for (int by = 0; by < comp.HB; by++)
                for (int bx = 0; bx < comp.WB; bx++)
                {
                    int baseIdx = (by * comp.Bpl + bx) * 64;
                    if (Ah == 0) // first pass
                    {
                        if (eobrun > 0) { eobrun--; continue; }
                        int k = Ss;
                        while (k <= Se)
                        {
                            int rs = Decode(br, act), r = rs >> 4, s = rs & 0xF;
                            if (s == 0)
                            {
                                if (r < 15) { eobrun = (1 << r) + (r != 0 ? br.Bits(r) : 0) - 1; break; }
                                k += 16;
                            }
                            else { k += r; comp.Coef[baseIdx + k] = Extend(br.Bits(s), s) << Al; k++; }
                        }
                    }
                    else // refinement pass
                    {
                        int k = Ss;
                        if (eobrun == 0)
                        {
                            while (k <= Se)
                            {
                                int rs = Decode(br, act), r = rs >> 4, s = rs & 0xF;
                                if (s != 0) { s = br.Bit() != 0 ? p1 : m1; }
                                else if (r != 15) { eobrun = (1 << r) + (r != 0 ? br.Bits(r) : 0); break; }

                                while (true)
                                {
                                    int ci = baseIdx + k;
                                    if (comp.Coef[ci] != 0)
                                    {
                                        if (br.Bit() != 0 && (comp.Coef[ci] & p1) == 0)
                                            comp.Coef[ci] += comp.Coef[ci] > 0 ? p1 : m1;
                                    }
                                    else { r--; if (r < 0) break; }
                                    k++; if (k > Se) break;
                                }
                                if (s != 0 && k <= Se) comp.Coef[baseIdx + k] = s;
                                k++;
                            }
                        }
                        if (eobrun > 0)
                        {
                            while (k <= Se)
                            {
                                int ci = baseIdx + k;
                                if (comp.Coef[ci] != 0 && br.Bit() != 0 && (comp.Coef[ci] & p1) == 0)
                                    comp.Coef[ci] += comp.Coef[ci] > 0 ? p1 : m1;
                                k++;
                            }
                            eobrun--;
                        }
                    }
                }
        }

        private sealed class BitW
        {
            public readonly List<byte> Out = new List<byte>();
            private int acc, cnt;
            public void Write(int code, int len)
            {
                for (int b = len - 1; b >= 0; b--)
                {
                    acc = (acc << 1) | ((code >> b) & 1);
                    if (++cnt == 8) { byte v = (byte)acc; Out.Add(v); if (v == 0xFF) Out.Add(0x00); acc = 0; cnt = 0; }
                }
            }
            public void Flush() { if (cnt > 0) Write(0x7F, 8 - cnt); }
        }

        private static int Cat(int x) { x = x < 0 ? -x : x; int n = 0; while (x != 0) { n++; x >>= 1; } return n; }
        private static int Mag(int v, int s) { if (v < 0) v = v - 1 + (1 << s); return v & ((1 << s) - 1); }

        private static byte[] EncodeBaseline(List<Comp> comps, List<byte[]> dqtSegs, int W, int H, int mcuW, int mcuH)
        {
            var dcL = HEnc.Build(DcLumBits, DcVals); var dcC = HEnc.Build(DcChrBits, DcVals);
            var acL = HEnc.Build(AcLumBits, AcLumVals); var acC = HEnc.Build(AcChrBits, AcChrVals);

            var o = new List<byte> { 0xFF, 0xD8 };
            foreach (var s in dqtSegs) o.AddRange(s);

            // SOF0
            var sof = new List<byte> { 8, (byte)(H >> 8), (byte)H, (byte)(W >> 8), (byte)W, (byte)comps.Count };
            foreach (var c in comps) { sof.Add((byte)c.Id); sof.Add((byte)((c.H << 4) | c.V)); sof.Add((byte)c.Q); }
            o.Add(0xFF); o.Add(0xC0); o.Add((byte)((sof.Count + 2) >> 8)); o.Add((byte)(sof.Count + 2)); o.AddRange(sof);

            // DHT (standard tables)
            void Dht(int tc, int th, byte[] bits1to16, byte[] vals)
            {
                int len = 2 + 1 + 16 + vals.Length;
                o.Add(0xFF); o.Add(0xC4); o.Add((byte)(len >> 8)); o.Add((byte)len);
                o.Add((byte)((tc << 4) | th));
                for (int l = 1; l <= 16; l++) o.Add(bits1to16[l]);
                o.AddRange(vals);
            }
            Dht(0, 0, DcLumBits, DcVals); Dht(0, 1, DcChrBits, DcVals);
            Dht(1, 0, AcLumBits, AcLumVals); Dht(1, 1, AcChrBits, AcChrVals);

            // APP14 Adobe, transform = 1 (YCbCr) so a viewer keeps the YCbCr interpretation
            o.Add(0xFF); o.Add(0xEE); o.Add(0x00); o.Add(0x0E);
            o.AddRange(new byte[] { (byte)'A', (byte)'d', (byte)'o', (byte)'b', (byte)'e' });
            o.AddRange(new byte[] { 0, 100, 0, 0, 0, 0, 1 });

            // SOS (baseline, interleaved, all components)
            var sos = new List<byte> { (byte)comps.Count };
            for (int ci = 0; ci < comps.Count; ci++) { int t = ci == 0 ? 0 : 1; sos.Add((byte)comps[ci].Id); sos.Add((byte)((t << 4) | t)); }
            sos.Add(0); sos.Add(63); sos.Add(0);
            o.Add(0xFF); o.Add(0xDA); o.Add((byte)((sos.Count + 2) >> 8)); o.Add((byte)(sos.Count + 2)); o.AddRange(sos);

            var bw = new BitW();
            var pred = new int[comps.Count];
            for (int my = 0; my < mcuH; my++)
                for (int mx = 0; mx < mcuW; mx++)
                    for (int ci = 0; ci < comps.Count; ci++)
                    {
                        var c = comps[ci];
                        bool chroma = ci != 0;
                        var dct = chroma ? dcC : dcL;
                        var act = chroma ? acC : acL;
                        for (int v = 0; v < c.V; v++)
                            for (int h = 0; h < c.H; h++)
                            {
                                int baseIdx = ((my * c.V + v) * c.Bpl + (mx * c.H + h)) * 64;
                                int dcVal = chroma ? 0 : c.Coef[baseIdx];
                                int diff = dcVal - pred[ci]; pred[ci] = dcVal;
                                int s = Cat(diff);
                                bw.Write(dct.Code[s], dct.Len[s]);
                                if (s != 0) bw.Write(Mag(diff, s), s);

                                if (chroma) { bw.Write(act.Code[0x00], act.Len[0x00]); } // EOB — all AC zero
                                else
                                {
                                    int run = 0;
                                    for (int k = 1; k < 64; k++)
                                    {
                                        int cf = c.Coef[baseIdx + k];
                                        if (cf == 0) run++;
                                        else
                                        {
                                            while (run > 15) { bw.Write(act.Code[0xF0], act.Len[0xF0]); run -= 16; }
                                            int sc = Cat(cf);
                                            int sym = (run << 4) | sc;
                                            bw.Write(act.Code[sym], act.Len[sym]);
                                            bw.Write(Mag(cf, sc), sc);
                                            run = 0;
                                        }
                                    }
                                    if (run > 0) bw.Write(act.Code[0x00], act.Len[0x00]);
                                }
                            }
                    }
            bw.Flush();
            o.AddRange(bw.Out);
            o.Add(0xFF); o.Add(0xD9);
            return o.ToArray();
        }
    }
}
