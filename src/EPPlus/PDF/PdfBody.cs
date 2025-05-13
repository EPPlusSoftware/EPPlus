using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF
{
    internal class PdfBody
    {
        internal int objectNumber;
        internal int version;
        internal string start = "obj\n";
        internal string end = "\nendobj\n";

        public PdfBody(int objectNumber, int version)
        {
            this.objectNumber = objectNumber;
            this.version = version;
        }
    }
}
