using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfFunctions
{

    //    Implemented    Function type
    //    [ ]            0 Sampled function
    //    [X]            2 Exponential interpolation function
    //    [ ]            3 Stitching function
    //    [ ]            4 PostScript calculator function

    internal abstract class PdfFunction : PdfObject
    {
        internal double[] Domain;
        internal double[] Range;

        public PdfFunction(int objectNumber, int version = 0) : base(objectNumber, version) { }
    }
}
