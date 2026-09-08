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
using EPPlus.Fonts.OpenType;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Fonts
{
    internal enum CIDFontSubtype
    {
        CIDFontType0,
        CIDFontType2
    }

    internal class PdfCIDFont : PdfObject
    {
        private readonly CIDFontSubtype Subtype;
        private readonly string BaseFont;
        private readonly CIDSystemInfo CIDInfoObject;
        private readonly int FontDescriptorObjectNumber;

        private int? DW;                                        // Default width
        private readonly List<object> W;                        // Width array
        private readonly int[] DW2;                             // Default metrics for vertical writing (2 numbers)
        private readonly List<object> W2;                       // Vertical writing metrics
        private readonly string CIDToGIDMap;                    // Can be string "Identity" or stream reference

        private HashSet<ushort> Gids;
        OpenTypeFont FontData;

        public PdfCIDFont(int objectNumber, OpenTypeFont fontData, HashSet<ushort> gids, CIDFontSubtype subtype, CIDSystemInfo CIDSystemInfoObject, string CIDToGDI, int fontDescriptorObjectNumber, int version = 0)
            : base(objectNumber, version)
        {
            Subtype = subtype;
            BaseFont = string.Concat(fontData.FullName.Where(c => !char.IsWhiteSpace(c)));
            CIDInfoObject = CIDSystemInfoObject;
            FontDescriptorObjectNumber = fontDescriptorObjectNumber;

            FontData = fontData;
            Gids = gids;

            DW = (int)Math.Round(1000.0d * 1000.0d / FontData.HeadTable.UnitsPerEm);//dw;
            CIDToGIDMap = CIDToGDI;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<<  /Type /Font\n" +
                            $"    /Subtype /{Subtype.ToString()}\n" +
                            $"    /BaseFont /{BaseFont}\n" +
                            $"    /CIDSystemInfo << /Registry ({CIDInfoObject.Registry}) /Ordering ({CIDInfoObject.Ordering}) /Supplement {CIDInfoObject.Supplement} >>\n" +
                            $"    /FontDescriptor {FontDescriptorObjectNumber.ToPdfStringF0()} 0 R");
            if (DW != null)
            {
                sb.AppendFormat($"\n    /DW {DW.ToPdfStringF0()}");
            }
            if (Gids != null)
            {
                sb.AppendFormat($"\n    /W [{BuildWidthsArray()}]");
            }
            if (DW2 != null)
            {
                sb.AppendFormat($"\n    /DW2 {DW2}");
            }
            if (W2 != null)
            {
                var widthsStr = string.Join(" ", W2.Select(w => w.ToString()).ToArray());
                sb.AppendFormat($"\n    /W2 [{widthsStr}]");
            }
            if (Subtype == CIDFontSubtype.CIDFontType2)
            {
                if(string.IsNullOrEmpty( CIDToGIDMap ))
                    sb.AppendFormat($"\n    /CIDToGIDMap /Identity");
                else
                    sb.AppendFormat($"\n    /CIDToGIDMap {CIDToGIDMap}");
            }
            sb.Append(" >>");
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<<  /Type /Font\n" +
                            $"    /Subtype /{Subtype.ToString()}\n" +
                            $"    /BaseFont /{BaseFont}\n" +
                            $"    /CIDSystemInfo << /Registry ({CIDInfoObject.Registry}) /Ordering ({CIDInfoObject.Ordering}) /Supplement {CIDInfoObject.Supplement} >>\n" +
                            $"    /FontDescriptor {FontDescriptorObjectNumber.ToPdfStringF0()} 0 R");
            if (DW != null)
            {
                sb.AppendFormat($"\n    /DW {DW.ToPdfStringF0()}");
            }
            if (Gids != null)
            {
                sb.AppendFormat($"\n    /W [ {BuildWidthsArray()} ]");
            }
            if (DW2 != null)
            {
                sb.AppendFormat($"\n    /DW2 {DW2}");
            }
            if (W2 != null)
            {
                var widthsStr = string.Join(" ", W2.Select(w => w.ToString()).ToArray());
                sb.AppendFormat($"\n    /W2 [ {widthsStr} ]");
            }
            if (Subtype == CIDFontSubtype.CIDFontType2)
            {
                if (string.IsNullOrEmpty(CIDToGIDMap))
                    sb.AppendFormat($"\n    /CIDToGIDMap /Identity");
                else
                    sb.AppendFormat($"\n    /CIDToGIDMap {CIDToGIDMap}");
            }
            sb.Append(" >>");
            WriteAscii(bw, sb.ToString());
        }

        private string BuildWidthsArray()
        {
            var sortedGids = Gids.OrderBy(g => g).ToList();
            var sb = new StringBuilder();
            int i = 0;
            while (i < sortedGids.Count)
            {
                ushort startGid = sortedGids[i];
                var widths = new List<int>();

                while (i < sortedGids.Count && sortedGids[i] == startGid + widths.Count)
                {
                    int rawWidth = FontData.HmtxTable.GetAdvanceWidth(sortedGids[i]);
                    int scaledWidth = (int)Math.Round(1000.0d * rawWidth / FontData.HeadTable.UnitsPerEm);
                    widths.Add(scaledWidth);
                    i++;
                }
                sb.Append($"{startGid} [");
                sb.Append(string.Join(" ", widths.Select(w => w.ToPdfStringF0()).ToArray()));
                sb.Append("] ");
            }
            return sb.ToString();
        }
    }
}