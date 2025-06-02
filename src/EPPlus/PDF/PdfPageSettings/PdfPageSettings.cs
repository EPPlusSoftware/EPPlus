using OfficeOpenXml.PDF.PdfPageSettings;
using OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfPageSettings
{
    public enum Orientations
    {
        Portrait,
        Landscape,
    }

    public enum GridLineTypes
    {
        Solid,
        Dotted,
        Lines,
    }

    public class PdfPageSettings
    {
        PdfPageSize _pageSize = PdfPageSize.A4;
        public PdfPageSize PageSize
        {
            get
            { 
                return _pageSize;
            }
            set
            {
                _pageSize = new PdfPageSize(value.Height, value.Width);
                if (value.Height > value.Width)
                {
                    _orientation = Orientations.Portrait;
                }
                else
                {
                    _orientation = Orientations.Landscape;
                }
            }
        }

        private Orientations _orientation = Orientations.Portrait;
        public Orientations Orientation
        {
            get
            {
                return _orientation;
            }
            set
            {
                if (value == Orientations.Portrait)
                {
                    if ((_pageSize.Height < _pageSize.Width))
                    {
                        _pageSize = new PdfPageSize(_pageSize.Height, _pageSize.Width);
                        _orientation = value;
                    }
                }
                if (value == Orientations.Landscape)
                {
                    if ((_pageSize.Height > _pageSize.Width))
                    {
                        _pageSize = new PdfPageSize(_pageSize.Height, _pageSize.Width);
                        _orientation = value;
                    }
                }
            }
        }

        public PdfMargins Margins = PdfMargins.Normal;

        public bool ShowGridLines = false;
        public GridLineTypes GridLineType = GridLineTypes.Solid;

        public bool ShowHeadings = false;



        internal PdfContentBounds ContentBounds;
        internal bool Debug = false;

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

//Page Breaks
    //Array of row and cols where to break to new page

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