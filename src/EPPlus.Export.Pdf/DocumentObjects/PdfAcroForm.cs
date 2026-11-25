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
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    // Minimal AcroForm whose only job is to declare that the document already
    // has form fields (see PdfFormFieldObject). NeedAppearances is false so no
    // viewer tries to regenerate appearances for the invisible field.
    internal class PdfAcroForm : PdfObject
    {
        private readonly List<int> fieldObjectNumbers;

        public PdfAcroForm(int objectNumber, List<int> fieldObjectNumbers, int version = 0)
            : base(objectNumber, version)
        {
            this.fieldObjectNumbers = fieldObjectNumbers;
        }

        private string Body()
        {
            var fields = string.Join(" ", fieldObjectNumbers.Select(n => $"{n.ToPdfStringF0()} 0 R"));
            return $"<< /Fields [ {fields} ]\n   /NeedAppearances false >>";
        }

        protected override void WriteBody(PdfWriter w) => w.Text(Body());
    }
}