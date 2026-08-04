/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  1/4/2021         EPPlus Software AB           Added wrapped text autofit mode
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// Determines how cells with WrapText enabled are measured when calculating
    /// column width in AutoFitColumns.
    /// </summary>
    public enum eWrappedTextAutofitMode
    {
        /// <summary>
        /// Cells with WrapText enabled are ignored when calculating column width.
        /// This is the default and matches the behaviour of earlier versions.
        /// </summary>
        Skip = 0,
        /// <summary>
        /// The entire cell text is measured as a single line, ignoring wrapping.
        /// </summary>
        FullText = 1,
        /// <summary>
        /// The text is split on explicit line breaks (CR, LF, CRLF) and the width
        /// of the widest resulting line determines the cell's contribution to the
        /// column width.
        /// </summary>
        SplitNewLine = 2,
        /// <summary>
        /// The text is split on whitespace (space, tab, and line breaks) and on
        /// hyphen characters (U+002D hyphen-minus and U+2010 hyphen). The width of
        /// the widest resulting segment determines the cell's contribution to the
        /// column width. Hyphens are visible and their width is included in the
        /// segment they terminate; whitespace is not. Note: this yields the minimum
        /// column width at which no single word overflows, and is not a simulation
        /// of Excel's line wrapping.
        /// </summary>
        SplitWord = 3
    }
}
