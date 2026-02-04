using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal class PdfCIDFont : PdfObject
    {
        public PdfCIDFont(int objectNumber, int version = 0)
            : base(objectNumber, version)
        {

        }

        internal override string RenderDictionary()
        {
            throw new NotImplementedException();
        }
    }
}
