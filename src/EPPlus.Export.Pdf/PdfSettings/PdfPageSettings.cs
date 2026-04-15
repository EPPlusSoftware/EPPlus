/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.PdfSettings.PdfPageSizes;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfSettings
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
        /// If true, subsetted fonts will be embedded into the PDF document.
        /// </summary>
        public bool EmbeddFonts = true;

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

        /// <summary>
        /// Set if comments and notes should be included.
        /// </summary>
        public CommentsAndNotes CommentsAndNotes = CommentsAndNotes.None;

        /// <summary>
        /// Set the starting page number.
        /// </summary>
        public int FirstPageNumber = 1;

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
                    if (_pageSize.Height < _pageSize.Width)
                    {
                        _pageSize = new PdfPageSize(_pageSize.Height, _pageSize.Width);
                        _orientation = value;
                    }
                }
                if (value == Orientations.Landscape)
                {
                    if (_pageSize.Height > _pageSize.Width)
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

    public enum CommentsAndNotes
    {
        /// <summary>
        /// Comments and Notes will be ignored.
        /// </summary>
        None,
        /// <summary>
        /// Comments and Notes will be displayed on a seprate page at the end.
        /// </summary>
        AtEndOfSheet,
        /// <summary>
        /// Notes will be displayed on the sheet. (Comments will not be shown.)
        /// </summary>
        AsDisplayedOnSheet
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