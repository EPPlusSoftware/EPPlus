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
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.Settings
{
    /// <summary>
    /// Settings object for exporting to PDF.
    /// </summary>
    internal class PdfRenderSettings
    {
        internal string Author { get; set; }
        internal string Title { get; set; }
        internal string LicenseType { get; set; }
        internal string LicenseHolder { get; set; }

        private OpenTypeFontEngine _fontEngine;
        /// <summary>
        /// Get the current font engine that is being used by EPPlus.
        /// </summary>
        internal OpenTypeFontEngine FontEngine
        {
            get
            {
                return _fontEngine;
            }
        }

        /// <summary>
        /// Add additional folders to search for fonts.
        /// </summary>
        internal List<string> FontDirectories = new List<string>();

        /// <summary>
        /// If true, epplus will look for fonts in the system directories for installed fonts.  c:\windows\fonts
        /// </summary>
        internal bool SearchSystemDirectories = true;

        /// <summary>
        /// If true, subsetted fonts will be embedded into the PDF document.
        /// </summary>
        internal bool EmbeddFonts = true;

        /// <summary>
        /// Prefix for marking a defined name as a bookmark for PDF export. The prefix gets stripped from the bookmark name.
        /// The name used for the bookmark is first check on the cells content. If the cell has no content it uses the defined name instead.
        /// </summary>
        internal string BookmarkNamePrefix = "bm_";

        /// <summary>
        /// Set true to supress Acrobat Readers auto detecting of form fillable areas.
        /// </summary>
        public bool SuppressFormFieldDetection = true;

        /// <summary>
        /// The order in how to create pages.
        /// </summary>
        internal PageOrders PageOrders = PageOrders.DownThenOver;

        /// <summary>
        /// Set to true to center content on page vertically.
        /// </summary>
        internal bool CenterOnPageVertically;

        /// <summary>
        /// Set to true to center content on page horizontally.
        /// </summary>
        internal bool CenterOnPageHorizontally;

        /// <summary>
        /// Set true to show grid lines.
        /// </summary>
        internal bool ShowGridLines = false;

        /// <summary>
        /// Set true to show row and column headings.
        /// </summary>
        internal bool ShowHeadings = false;

        /// <summary>
        /// Set to true for to only use black and white.
        /// </summary>
        internal bool BlackAndWhite = false;

        /// <summary>
        /// Set to true for to make a draft.
        /// </summary>
        internal bool Draft = false;

        /// <summary>
        /// Set the range to repeat at the top of the page.
        /// </summary>
        internal string RowsToRepeatAtTop = null;

        /// <summary>
        /// Set the range to repeat to the left of the page.
        /// </summary>
        internal string ColumnsToRepeatAtLeft = null;

        /// <summary>
        /// Specific range to print.
        /// </summary>
        internal string PrintArea = null;

        /// <summary>
        /// Set if comments and notes should be included.
        /// </summary>
        internal CommentsAndNotes CommentsAndNotes = CommentsAndNotes.None;

        /// <summary>
        /// Sets how to display errors in cells.
        /// </summary>
        internal CellErrors CellErrors = CellErrors.Displayed;

        /// <summary>
        /// Set the starting page number.
        /// </summary>
        internal int FirstPageNumber = 1;

        PdfPageSize _pageSize = PdfPageSize.A4;
        /// <summary>
        /// Set the size of pages.
        /// </summary>
        internal PdfPageSize PageSize
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
        internal Orientations Orientation
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
        internal PdfMargins Margins
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
        internal PdfScaling Scaling
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

        private double _resolvedFitScale = 1d;
        internal void SetResolvedFitScale(double s) { _resolvedFitScale = s; }

        internal bool ScaleHeaderWithDocument { get; set; }

        internal double ContentScale
        {
            get
            {
                if (_scaling == null) 
                   return 1d;
                if (_scaling.ScalingMode == ScalingMode.AdjustToNormalSize)
                    return _scaling.Scale > 0d ? _scaling.Scale : 1d;
                return _resolvedFitScale;
            }
        }
        internal double EffectiveContentWidth => ContentBounds.Width / ContentScale;
        internal double EffectiveContentHeight => ContentBounds.Height / ContentScale;

        /// <summary>
        /// GSUB (glyph substitution) features to request when shaping text for this export, e.g.
        /// ligatures and contextual alternates. Defaults to <see cref="GsubFeature.Liga"/> |
        /// <see cref="GsubFeature.Clig"/> | <see cref="GsubFeature.Calt"/> | <see cref="GsubFeature.Ccmp"/>.
        /// </summary>
        /// <remarks>
        /// This setting applies to the whole export - there is currently no per-cell or
        /// per-range override. Whether a requested feature has any visible effect still depends
        /// on the font actually defining it; see <see cref="GsubFeature"/> for what each flag
        /// means and which fonts typically support it.
        /// </remarks>
        internal GsubFeature GsubFeatures { get; set; } = GsubFeature.Liga | GsubFeature.Clig | GsubFeature.Calt | GsubFeature.Ccmp;

        /// <summary>
        /// GPOS (glyph positioning) features to request when shaping text for this export, e.g.
        /// kerning and mark attachment. Defaults to <see cref="GposFeature.Kern"/> | <see cref="GposFeature.Mark"/>.
        /// </summary>
        /// <remarks>
        /// This setting applies to the whole export - there is currently no per-cell or
        /// per-range override. Whether a requested feature has any visible effect still depends
        /// on the font actually defining it; see <see cref="GposFeature"/> for what each flag
        /// means.
        /// </remarks>
        internal GposFeature GposFeatures { get; set; } = GposFeature.Kern | GposFeature.Mark;

        internal PdfContentBounds ContentBounds = new PdfContentBounds(PdfMargins.Normal, PdfPageSize.A4);
        internal string defaultFontName = "";

        internal string AdditionalContent;

        //DEBUG
        internal bool Debug = false;
        internal bool PrintAsText = false;

        internal PdfRenderSettings(OpenTypeFontEngine fontEngine)
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

        internal PdfRenderSettings CloneForSheet()
        {
            var c = (PdfRenderSettings)MemberwiseClone();
            c.ContentBounds = new PdfContentBounds(_margins, _pageSize);
            return c;
        }
    }

    /// <summary>
    /// Orientation of pages.
    /// </summary>
    internal enum Orientations
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
    internal enum PageOrders
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
    internal enum CommentsAndNotes
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
    internal enum CellErrors
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