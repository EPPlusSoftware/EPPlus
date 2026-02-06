using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal struct CIDSystemInfo
    {
        public string Registry { get; set; }
        public string Ordering { get; set; }
        public int Supplement { get; set; }
    }
}
