using OfficeOpenXml.Encryption;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    public enum CIDFontSubtype
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

        private readonly int? DW;                               // Default width
        private readonly List<object> W;                        // Width array
        private readonly int[] DW2;                             // Default metrics for vertical writing (2 numbers)
        private readonly List<object> W2;                       // Vertical writing metrics
        private readonly object CIDToGIDMap;                    // Can be string "Identity" or stream reference

        public PdfCIDFont(int objectNumber, CIDFontSubtype subtype, string baseFont, CIDSystemInfo CIDSystemInfoObject, int fontDescriptorObjectNumber, int? dw = null, List<object> w = null, int[] dw2 = null, List<object> w2 = null, object CIDToGDI = null, int version = 0)
            : base(objectNumber, version)
        {
            Subtype = subtype;
            BaseFont = baseFont;
            CIDInfoObject = CIDSystemInfoObject;
            FontDescriptorObjectNumber = fontDescriptorObjectNumber;
            DW = dw;
            W = w;
            DW2 = dw2;
            W2 = w2;
            CIDToGIDMap = CIDToGDI;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<<  /Type /Font\n" +
                            $"    /SubType /{Subtype.ToString()}\n" +
                            $"    /BaseFont /{BaseFont}\n" +
                            $"    /CIDSystemInfo << /Registry ({CIDInfoObject.Registry}) /Ordering ({CIDInfoObject.Ordering}) /Supplement ({CIDInfoObject.Supplement}) >>\n" +
                            $"    /FontDescriptor {FontDescriptorObjectNumber} 0 R");
            if (DW != null)
            {
                sb.AppendFormat($"\n    /DW {DW}");
            }
            if (W != null)
            {
                var widthsStr = string.Join(" ", W.Select(w => w.ToString()).ToArray());
                sb.AppendFormat($"\n    /W [{widthsStr}]");
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
                if(CIDToGIDMap != null)
                    sb.AppendFormat($"\n    /CIDToGIDMap {CIDToGIDMap.ToString()}");
                else
                    sb.AppendFormat($"\n    /CIDToGIDMap /Identity");
            }
            sb.Append(" >>");
            return sb.ToString();
        }
    }
}