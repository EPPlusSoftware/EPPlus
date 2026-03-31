using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfHeaderFooterCollection
    {
        //H = Header  F = Footer
        //1 = First   O = Odd     E = Even
        //L = Left    C = Center  R = Right

        //   H  H  H
        //1 [L][C][R]
        //O [L][C][R]
        //E [L][C][R]

        //   F  F  F
        //1 [L][C][R]
        //O [L][C][R]
        //E [L][C][R]

        public PdfHeaderFooter[,,] pdfHeaderFooters = new PdfHeaderFooter[3, 3, 2];;
        public PdfHeaderFooterCollection(ExcelHeaderFooter headerFooter)
        {
            pdfHeaderFooters 
        }
    }
}
