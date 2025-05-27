using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfFontWidths : PdfObject
    {
        internal readonly int[] widths;
        internal readonly int firstChar;
        internal readonly int lastChar;

        public PdfFontWidths(int objectNumber, int[] widths, int firstChar, int lastChar, int version = 0)
            : base(objectNumber, version)
        {
            this.widths = widths;
            this.firstChar = firstChar;
            this.lastChar = lastChar;
        }

        internal override string RenderDictionary()
        {
            throw new NotImplementedException();
        }
    }
}
