using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal class PdfUnicodeMap : PdfObject
    {
        public PdfUnicodeMap(int objectNumber, int version = 0)
            : base(objectNumber, version)
        {
        }

        internal override string RenderDictionary()
        {
            throw new NotImplementedException();
        }
    }
}
