using OfficeOpenXml.PDF.PdfSettings.PdfPageSizes;
using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfSettings
{
    /// <summary>
    /// Settings object for exporting to PDF.
    /// </summary>
    public class PdfPageSettings
    {
        /// <summary>
        /// Add additional folders to search for fonts.
        /// </summary>
        public List<string> FontDirectories = new List<string>();

        /// <summary>
        /// If true, epplus will look for fonts in the system directories for installed fonts.  c:\windows\fonts
        /// </summary>
        public bool SearchSystemDirectories = true;

        /// <summary>
        /// The order in how to create pages.
        /// </summary>
        public PageOrders PageOrders = PageOrders.DownThenOver;

        /// <summary>
        /// Set to true to center content on page vertically.
        /// </summary>
        public bool CenterOnPageVertically;

        /// <summary>
        /// Set to true to center content on page horizontally.
        /// </summary>
        public bool CenterOnPageHorizontally;

        /// <summary>
        /// Set true to show grid lines.
        /// </summary>
        public bool ShowGridLines = false;

        /// <summary>
        /// Set true to show row and column headings.
        /// </summary>
        public bool ShowHeadings = false;

        PdfPageSize _pageSize = PdfPageSize.A4;
        /// <summary>
        /// Set the size of pages.
        /// </summary>
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
                ContentBounds.CalculateBounds(Margins, PageSize);
            }
        }

        private Orientations _orientation = Orientations.Portrait;
        /// <summary>
        /// Set the orientation of the pages.
        /// </summary>
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

        private PdfMargins _margins = PdfMargins.Normal;
        /// <summary>
        /// Set the page margins.
        /// </summary>
        public PdfMargins Margins
        {
            get
            {
                return _margins;
            }
            set
            {
                _margins = value;
                ContentBounds.CalculateBounds(Margins, PageSize);
            }
        }

        internal PdfContentBounds ContentBounds = new PdfContentBounds(PdfMargins.Normal, PdfPageSize.A4);
        internal string defaultFontName = "";

        //DEBUG
        internal bool Debug = false;
        internal bool PrintAsText = false;
    }

    /// <summary>
    /// Orientation of pages.
    /// </summary>
    public enum Orientations
    {
        /// <summary>
        /// Portrait orientation.
        /// </summary>
        Portrait,
        /// <summary>
        /// Landscape orientation.
        /// </summary>
        Landscape,
    }

    /// <summary>
    /// Order of pages.
    /// </summary>
    public enum PageOrders
    {
        /// <summary>
        /// Order Down then over.
        /// </summary>
        DownThenOver,
        /// <summary>
        /// Order Over then down.
        /// </summary>
        OverThenDown,
    }
}


/*
//Page
    //Orientation
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