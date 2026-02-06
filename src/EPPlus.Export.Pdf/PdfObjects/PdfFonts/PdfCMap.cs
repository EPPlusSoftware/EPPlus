using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal class PdfCMap : PdfObject
    {
        private readonly string CMapName;
        private readonly CIDSystemInfo CIDInfoObject;

        private readonly int WMode;
        private readonly string UseCMap;

        public PdfCMap(int objectNumber, string CmapName, CIDSystemInfo CIDSystemInfoObject, int WMode = -1, string UseCMap = "", int version = 0) : base(objectNumber, version)
        {
            this.CMapName = CmapName;
            this.CIDInfoObject = CIDSystemInfoObject;
            this.WMode = WMode;
            this.UseCMap = UseCMap;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            //sb.AppendFormat($"<<  /Type /CMap\n" +
            //                $"    /CMapName /{CMapName}\n" +
            //                $"    /CIDSystemInfo << /Registry ({CIDInfoObject.Registry}) /Ordering ({CIDInfoObject.Ordering}) /Supplement ({CIDInfoObject.Supplement}) >>\n" +

            //if (DW != null)
            //{
            //    sb.AppendFormat($"\n    /DW {DW}");
            //}
            //sb.Append(" >>");
            return sb.ToString();
        }
    }
}
