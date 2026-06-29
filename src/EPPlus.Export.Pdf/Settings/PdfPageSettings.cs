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
                if(_fontEngine == null)
                {
                    _fontEngine = new OpenTypeFontEngine(x =>
                    {
                        if(FontDirectories != null && FontDirectories.Any())
                        {
                            foreach(var dir in FontDirectories)
                            {
                                if (!System.IO.Directory.Exists(dir))
                                {
                                    throw new System.IO.DirectoryNotFoundException($"Font directory not found: {dir}");
                                }
                                x.FontDirectories.Add(dir);
                            }
                            x.SearchSystemDirectories = SearchSystemDirectories;

                        }
                        
                    });
                }
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
        /// Set the range to repeat at the top of the page.
        /// </summary>
        public string RowsToRepeatAtTop = null;

        /// <summary>
        /// Set the range to repeat to the left of the page.
        /// </summary>
        public string ColumnsToRepeatAtLeft = null;

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
        /// Comments and Notes will be displayed on a seprate page at the end.
        /// </summary>
        AtEndOfSheet,
        /// <summary>
        /// Notes will be displayed on the sheet. (Comments will not be shown.)
        /// </summary>
        AsDisplayedOnSheet
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