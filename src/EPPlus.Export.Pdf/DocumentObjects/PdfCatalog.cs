/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;
using System.IO;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfCatalog : PdfObject
    {
        private readonly int pagesObjectNumber;
        internal int NamesObjectNumber = 0;
        internal int OutlinesObjectNumber = 0;

        public PdfCatalog(int objectNumber, int pagesObjectNumber, int version = 0)
            : base(objectNumber, version)
        {
            this.pagesObjectNumber = pagesObjectNumber;
        }

        private string Body()
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("<< /Type /Catalog\n");
            sb.Append($"   /Pages {pagesObjectNumber.ToPdfStringF0()} 0 R");
            if (NamesObjectNumber > 0)
                sb.Append($"\n   /Names {NamesObjectNumber.ToPdfStringF0()} 0 R");
            if (OutlinesObjectNumber > 0)
            {
                sb.Append($"\n   /Outlines {OutlinesObjectNumber.ToPdfStringF0()} 0 R");
                sb.Append("\n   /PageMode /UseOutlines");
            }
            sb.Append(" >>");
            return sb.ToString();
        }

        internal override string RenderDictionary() => Body();
        internal override void RenderDictionary(BinaryWriter bw) => WriteAscii(bw, Body());
    }
}
