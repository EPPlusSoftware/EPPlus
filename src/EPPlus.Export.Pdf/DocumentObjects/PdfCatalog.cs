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

        public PdfCatalog(int objectNumber, int pagesObjectNumber, int version = 0)
            : base(objectNumber, version)
        {
            this.pagesObjectNumber = pagesObjectNumber;
        }

        internal override string RenderDictionary()
        {
            return $"<< /Type /Catalog\n" +
                   $"   /Pages {pagesObjectNumber.ToPdfStringF0()} 0 R >>";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, $"<< /Type /Catalog\n" +
                           $"   /Pages {pagesObjectNumber.ToPdfStringF0()} 0 R >>");
        }
    }
}
