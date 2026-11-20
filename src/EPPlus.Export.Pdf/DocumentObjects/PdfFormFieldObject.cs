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
    // A single hidden, read-only widget/field. Its only purpose is to make the
    // document count as an already-prepared AcroForm, so Acrobat's Fill & Sign
    // auto-detection stands down instead of offering fillable areas over our
    // gridline tables. It never renders and cannot be interacted with.
    internal class PdfFormFieldObject : PdfObject
    {
        private readonly int pageObjectNumber;
        private readonly string fieldName;

        public PdfFormFieldObject(int objectNumber, int pageObjectNumber, string fieldName = "epplus_noautofield", int version = 0)
            : base(objectNumber, version)
        {
            this.pageObjectNumber = pageObjectNumber;
            this.fieldName = fieldName;
        }

        private string Body()
        {
            return $"<< /Type /Annot /Subtype /Widget\n" +
                   $"   /FT /Tx\n" +
                   $"   /T {fieldName.ToPdfTextString()}\n" +
                   $"   /Ff 1\n" +
                   $"   /F 2\n" +
                   $"   /Rect [ 0 0 1 1 ]\n" +
                   $"   /P {pageObjectNumber.ToPdfStringF0()} 0 R >>";
        }

        internal override string RenderDictionary() => Body();
        internal override void RenderDictionary(BinaryWriter bw) => WriteAscii(bw, Body());
    }
}