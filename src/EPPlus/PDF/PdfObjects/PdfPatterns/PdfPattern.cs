using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal abstract class PdfPattern : PdfObject
    {
        public PdfPattern(int objectNumber, int version = 0) : base(objectNumber, version) { }
    }
}
