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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Helpers
{
    internal static class JpegBlackWhiteHelper
    {
        internal static byte[] MakeGrayscale(byte[] jpeg)
        {
            try
            {
                if (FindSofMarker(jpeg) == 0xC2)                 // progressive
                    return JpegProgressiveHelper.ToGrayBaseline(jpeg) ?? jpeg;
                return TryTransform(jpeg) ?? jpeg;               // your existing baseline/YCCK transcoder
            }
            catch { return jpeg; }
        }

        private static int FindSofMarker(byte[] d)
        {
            if (d == null || d.Length < 4 || (d[0] & 0xFF) != 0xFF || (d[1] & 0xFF) != 0xD8) return -1;
            int i = 2;
            while (i + 1 < d.Length)
            {
                if ((d[i] & 0xFF) != 0xFF) { i++; continue; }
                int m = d[i + 1] & 0xFF;
                if (m == 0xFF) { i++; continue; }
                if (m == 0xD8 || m == 0xD9 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
                if (m >= 0xC0 && m <= 0xCF && m != 0xC4 && m != 0xC8 && m != 0xCC) return m;  // SOFn
                if (m == 0xDA) return -1;                        // reached the scan, no SOF found
                i += 2 + (((d[i + 2] & 0xFF) << 8) | (d[i + 3] & 0xFF));
            }
            return -1;
        }

        private sealed class Huff
        {
            public byte[] Bits = new byte[17];   // Bits[1..16] = count of codes of that length
            public byte[] Vals;                  // HUFFVAL
            public int[] MinCode = new int[17];
            public int[] MaxCode = new int[17];
            public int[] ValPtr = new int[17];
            public int[] EncCode = new int[256];
            public int[] EncLen = new int[256];

            public void Build()
            {
                int lastk = 0;
                var size = new int[257];
                for (int l = 1; l <= 16; l++)
                    for (int i = 0; i < Bits[l]; i++) size[lastk++] = l;
                size[lastk] = 0;

                var code = new int[257];
                int k = 0, c = 0, si = size[0];
                if (lastk > 0)
                {
                    while (true)
                    {
                        do { code[k] = c; c++; k++; } while (size[k] == si);
                        if (size[k] == 0) break;
                        do { c <<= 1; si++; } while (size[k] != si);
                    }
                }

                k = 0;
                for (int l = 1; l <= 16; l++)
                {
                    if (Bits[l] == 0) { MaxCode[l] = -1; }
                    else { ValPtr[l] = k; MinCode[l] = code[k]; k += Bits[l]; MaxCode[l] = code[k - 1]; }
                }
                for (int idx = 0; idx < lastk; idx++)
                {
                    EncCode[Vals[idx]] = code[idx];
                    EncLen[Vals[idx]] = size[idx];
                }
            }
        }

        private sealed class BitReader
        {
            private readonly byte[] d; private int pos; private readonly int end;
            private int buf, cnt;
            public BitReader(byte[] data, int start, int stop) { d = data; pos = start; end = stop; }

            public int Bit()
            {
                if (cnt == 0)
                {
                    if (pos >= end) return 0;
                    int b = d[pos++] & 0xFF;
                    if (b == 0xFF && pos < end && (d[pos] & 0xFF) == 0x00) pos++; // unstuff
                    buf = b; cnt = 8;
                }
                cnt--; return (buf >> cnt) & 1;
            }
            public int Bits(int n) { int v = 0; while (n-- > 0) v = (v << 1) | Bit(); return v; }

            public void RestartAlign()
            {
                cnt = 0;                                   // discard partial bits
                if (pos + 1 < end && (d[pos] & 0xFF) == 0xFF
                    && (d[pos + 1] & 0xFF) >= 0xD0 && (d[pos + 1] & 0xFF) <= 0xD7)
                    pos += 2;                              // consume RSTn
            }
        }

        private sealed class BitWriter
        {
            public readonly List<byte> Out = new List<byte>();
            private int acc, cnt;
            public void Write(int code, int len)
            {
                for (int i = len - 1; i >= 0; i--)
                {
                    acc = (acc << 1) | ((code >> i) & 1);
                    if (++cnt == 8)
                    {
                        byte v = (byte)acc; Out.Add(v);
                        if (v == 0xFF) Out.Add(0x00);      // stuff
                        acc = 0; cnt = 0;
                    }
                }
            }
            public void FlushToByte() { if (cnt > 0) Write(0x7F, 8 - cnt); }
            private int rst;
            public void EmitRst() { FlushToByte(); Out.Add(0xFF); Out.Add((byte)(0xD0 + (rst & 7))); rst++; }
        }

        private static byte[] TryTransform(byte[] d)
        {
            if (d == null || d.Length < 4 || (d[0] & 0xFF) != 0xFF || (d[1] & 0xFF) != 0xD8) return null;

            var dcT = new Huff[4]; var acT = new Huff[4];
            int comps = 0;
            var cid = new int[4]; var cH = new int[4]; var cV = new int[4];
            int W = 0, H = 0, Hmax = 1, Vmax = 1, dri = 0;
            int sosEnd = -1, scanTermMarker = -1;
            int sNs = 0; var sSof = new int[4]; var sTd = new int[4]; var sTa = new int[4];
            int adobeTransform = -1;

            int i = 2;
            while (i + 1 < d.Length)
            {
                if ((d[i] & 0xFF) != 0xFF) { i++; continue; }
                int m = d[i + 1] & 0xFF;
                if (m == 0xFF) { i++; continue; }
                if (m == 0xD9) break;
                if (m == 0xC2 || m == 0xC3 || (m >= 0xC5 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF))
                    return null;                                   // progressive / arithmetic / unsupported
                int seg = ((d[i + 2] & 0xFF) << 8) | (d[i + 3] & 0xFF);
                int p = i + 4;

                if (m == 0xC0 || m == 0xC1)                        // SOF0 / SOF1 (baseline / ext. sequential)
                {
                    H = ((d[p + 1] & 0xFF) << 8) | (d[p + 2] & 0xFF);
                    W = ((d[p + 3] & 0xFF) << 8) | (d[p + 4] & 0xFF);
                    comps = d[p + 5] & 0xFF;
                    if (comps != 3 && comps != 4) return null;
                    for (int cc = 0; cc < comps; cc++)
                    {
                        int q = p + 6 + cc * 3;
                        cid[cc] = d[q] & 0xFF;
                        cH[cc] = (d[q + 1] >> 4) & 0xF;
                        cV[cc] = d[q + 1] & 0xF;
                        if (cH[cc] > Hmax) Hmax = cH[cc];
                        if (cV[cc] > Vmax) Vmax = cV[cc];
                    }
                }
                else if (m == 0xEE && seg >= 14 &&
                         d[p] == (byte)'A' && d[p + 1] == (byte)'d' && d[p + 2] == (byte)'o' &&
                         d[p + 3] == (byte)'b' && d[p + 4] == (byte)'e')
                {
                    adobeTransform = d[p + 11] & 0xFF;
                }
                else if (m == 0xC4)                                // DHT (may hold several tables)
                {
                    int q = p, segEnd = i + 2 + seg;
                    while (q < segEnd)
                    {
                        int tc = (d[q] >> 4) & 0xF, th = d[q] & 0xF; q++;
                        var t = new Huff(); int tot = 0;
                        for (int l = 1; l <= 16; l++) { t.Bits[l] = d[q++]; tot += t.Bits[l]; }
                        t.Vals = new byte[tot];
                        for (int v = 0; v < tot; v++) t.Vals[v] = d[q++];
                        t.Build();
                        if (tc == 0) dcT[th] = t; else acT[th] = t;
                    }
                }
                else if (m == 0xDD) { dri = ((d[p] & 0xFF) << 8) | (d[p + 1] & 0xFF); }  // DRI
                else if (m == 0xDA)                                // SOS
                {
                    sNs = d[p] & 0xFF;
                    for (int s = 0; s < sNs; s++)
                    {
                        int selId = d[p + 1 + s * 2] & 0xFF;
                        int tdta = d[p + 2 + s * 2] & 0xFF;
                        int sof = 0; for (int cc = 0; cc < comps; cc++) if (cid[cc] == selId) sof = cc;
                        sSof[s] = sof; sTd[s] = (tdta >> 4) & 0xF; sTa[s] = tdta & 0xF;
                    }
                    sosEnd = p + 1 + sNs * 2 + 3;                  // after Ss, Se, Ah/Al
                    int j = sosEnd;
                    while (j + 1 < d.Length)
                    {
                        if ((d[j] & 0xFF) == 0xFF)
                        {
                            int mm = d[j + 1] & 0xFF;
                            if (mm != 0x00 && !(mm >= 0xD0 && mm <= 0xD7)) { scanTermMarker = j; break; }
                        }
                        j++;
                    }
                    break;
                }
                i += 2 + seg;
            }

            if (sosEnd < 0 || scanTermMarker < 0 || W == 0 || H == 0) return null;

            bool threeYcbcr = comps == 3 && sNs == 3 && adobeTransform != 0;
            bool fourYcck = comps == 4 && sNs == 4 && adobeTransform == 2;   //            transform 0 on 3 comps = RGB, leave alone
            if (!threeYcbcr && !fourYcck) return null;

            var br = new BitReader(d, sosEnd, scanTermMarker);
            var bw = new BitWriter();

            int mcuX = (W + 8 * Hmax - 1) / (8 * Hmax);
            int mcuY = (H + 8 * Vmax - 1) / (8 * Vmax);
            long totalMcu = (long)mcuX * mcuY;

            for (long mcu = 0; mcu < totalMcu; mcu++)
            {
                if (dri > 0 && mcu > 0 && mcu % dri == 0) { bw.EmitRst(); br.RestartAlign(); }

                for (int s = 0; s < sNs; s++)
                {
                    var dc = dcT[sTd[s]]; var ac = acT[sTa[s]];
                    if (dc == null || ac == null) return null;
                    bool passthrough = sSof[s] != 1 && sSof[s] != 2;
                    int blocks = cH[sSof[s]] * cV[sSof[s]];            //            zero only chroma (indices 1,2);
                                                                       //            keep Y (0) and, for YCCK, K (3)
                    for (int b = 0; b < blocks; b++)
                    {
                        // ---- DC ----
                        int t = Decode(br, dc);
                        int dcBits = br.Bits(t);
                        if (passthrough)
                        {
                            if (dc.EncLen[t] == 0) return null;
                            bw.Write(dc.EncCode[t], dc.EncLen[t]);
                            if (t > 0) bw.Write(dcBits, t);
                        }
                        // ---- AC ----
                        int k = 1;
                        while (k < 64)
                        {
                            int rs = Decode(br, ac);
                            int run = rs >> 4, size = rs & 0xF;
                            if (size == 0)
                            {
                                if (passthrough) { if (ac.EncLen[rs] == 0) return null; bw.Write(ac.EncCode[rs], ac.EncLen[rs]); }
                                if (run != 15) break;                   // EOB
                                k += 16; continue;                      // ZRL
                            }
                            int acBits = br.Bits(size);
                            if (passthrough)
                            {
                                if (ac.EncLen[rs] == 0) return null;
                                bw.Write(ac.EncCode[rs], ac.EncLen[rs]);
                                bw.Write(acBits, size);
                            }
                            k += run + 1;
                        }

                        if (!passthrough)
                        {
                            if (dc.EncLen[0] == 0 || ac.EncLen[0] == 0) return null;
                            bw.Write(dc.EncCode[0], dc.EncLen[0]);      // DC category 0 (diff 0)
                            bw.Write(ac.EncCode[0], ac.EncLen[0]);      // AC EOB
                        }
                    }
                }
            }

            bw.FlushToByte();

            var outp = new List<byte>(d.Length);
            for (int k = 0; k < sosEnd; k++) outp.Add(d[k]);            // headers + SOS header, verbatim
            outp.AddRange(bw.Out);                                      // re-encoded entropy stream
            for (int k = scanTermMarker; k < d.Length; k++) outp.Add(d[k]); // EOI + trailer, verbatim
            return outp.ToArray();
        }

        private static int Decode(BitReader br, Huff h)
        {
            int l = 1, code = br.Bit();
            while (h.MaxCode[l] < 0 || code > h.MaxCode[l])
            {
                code = (code << 1) | br.Bit();
                if (++l > 16) return 0;
            }
            return h.Vals[h.ValPtr[l] + code - h.MinCode[l]] & 0xFF;
        }
    }
}
