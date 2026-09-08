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
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.Settings
{
    /// <summary>
    /// Settings object for exporting to PDF.
    /// </summary>
    public class PdfPageSettings
    {
        private OpenTypeFontEngine _fontEngine;
        /// <summary>
        /// Get the current font engine that is being used by EPPlus.
        /// </summary>
        public OpenTypeFontEngine FontEngine
        {
            get
            {
                return _fontEngine;
            }
        }

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
        /// Set to true for to only use black and white.
        /// </summary>
        public bool BlackAndWhite = false;

        /// <summary>
        /// Set to true for to make a draft.
        /// </summary>
        public bool Draft = false;

        /// <summary>
        /// Set the range to repeat at the top of the page.
        /// </summary>
        public string RowsToRepeatAtTop = null;

        /// <summary>
        /// Set the range to repeat to the left of the page.
        /// </summary>
        public string ColumnsToRepeatAtLeft = null;

        /// <summary>
        /// Specific range to print.
        /// </summary>
        public string PrintArea = null;

        /// <summary>
        /// Set if comments and notes should be included.
        /// </summary>
        public CommentsAndNotes CommentsAndNotes = CommentsAndNotes.None;

        /// <summary>
        /// Sets how to display errors in cells.
        /// </summary>
        public CellErrors CellErrors = CellErrors.Displayed; 

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
                var size = new PdfPageSize(value.Width, value.Height); // store as authored, no transpose
                if (_orientationExplicitlySet)
                {
                    size = ApplyOrientation(size, _orientation);
                }
                else
                {
                    _orientation = size.Height >= size.Width
                        ? Orientations.Portrait
                        : Orientations.Landscape;
                }
                _pageSize = size;
                ContentBounds.CalculateBounds(Margins, _pageSize);
            }
        }

        private bool _orientationExplicitlySet = false;
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
                _orientation = value;
                _orientationExplicitlySet = true;
                _pageSize = ApplyOrientation(_pageSize, value);
                ContentBounds.CalculateBounds(Margins, _pageSize);
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

        private PdfScaling _scaling = PdfScaling.NoScaling;
        /// <summary>
        /// Set the scaling. NOT IMPLEMENTED
        /// </summary>
        public PdfScaling Scaling
        {
            get
            {
                return _scaling;
            }
            set
            {
                _scaling = value;
            }
        }

        internal PdfContentBounds ContentBounds = new PdfContentBounds(PdfMargins.Normal, PdfPageSize.A4);
        internal string defaultFontName = "";

        //DEBUG
        internal bool Debug = true;
        internal bool PrintAsText = true;

        public PdfPageSettings(OpenTypeFontEngine fontEngine)
        {
            _fontEngine = fontEngine;
        }

        private static PdfPageSize ApplyOrientation(PdfPageSize size, Orientations orientation)
        {
            bool isPortrait = size.Height >= size.Width;
            bool wantPortrait = orientation == Orientations.Portrait;
            return isPortrait == wantPortrait
                ? size
                : new PdfPageSize(size.Height, size.Width); // swap (ctor is width, height)
        }

        internal PdfPageSettings CloneForSheet()
        {
            var c = new PdfPageSettings(_fontEngine);
            c.FontDirectories = FontDirectories;
            c.SearchSystemDirectories = SearchSystemDirectories;
            c.EmbeddFonts = EmbeddFonts;
            c.defaultFontName = defaultFontName;
            c.Debug = Debug;
            c.PrintAsText = PrintAsText;
            return c;
        }
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

    /// <summary>
    /// How comments will be displayed.
    /// </summary>
    public enum CommentsAndNotes
    {
        /// <summary>
        /// Comments and Notes will be ignored.
        /// </summary>
        None,
        /// <summary>
        /// Notes will be displayed on the sheet. (Comments will not be shown.)
        /// </summary>
        AsDisplayedOnSheet,
        /// <summary>
        /// Comments and Notes will be displayed on a seprate page at the end.
        /// </summary>
        AtEndOfSheet,
    }

    /// <summary>
    /// How errors will be displayed.
    /// </summary>
    public enum CellErrors
    {
        /// <summary>
        /// Errors will be displayed
        /// </summary>
        Displayed,
        /// <summary>
        /// Errors will be not displayed
        /// </summary>
        Blank,
        /// <summary>
        /// Errors will be displayed as "--"
        /// </summary>
        Dashed,
        /// <summary>
        /// Errors will be displayed as #N/A
        /// </summary>
        NA,
    }
}