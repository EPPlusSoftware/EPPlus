using OfficeOpenXml.PDF.PdfPageSettings;
using OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfPageSettings
{
    public class PdfPageSettings
    {
        public PdfPageSize PageSize = PdfPageSize.A4;
        public PdfMargins Margins = new PdfMargins();

        internal PdfContentBounds ContentBounds;

    }
}


/*
//Page
    /Orientation
        //Portrait
        //Landscape
    //Scaling
    //First Page number

//marigns
    //Header
    //Footer
    //Center On Page
        //Horizontal
        //vertical

//Sheet
    //print grid lines
    //black and white
    //print cell errors
    //comments and notes
    //Row and column headings
    //Page order
    //down, then over
    //over, then down
*/