/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// Copy options enum. Specify the flags that you want to exclude from the copy or if you want to transpose the output.
    /// </summary>
    [Flags]    
    public enum ExcelRangeCopyOptionFlags : int
    {
        /// <summary>
        /// Exclude formulas from being copied. Only the value of the cell will be copied
        /// </summary>
        ExcludeFormulas = 0x1,
        /// <summary>
        /// Will exclude formulas and values from being copied
        /// </summary>
        ExcludeValues = 0x2,
        /// <summary>
        /// Exclude styles from being copied. 
        /// </summary>
        ExcludeStyles = 0x4,
        /// <summary>
        /// Exclude comments from being copied. 
        /// </summary>
        ExcludeComments = 0x8,
        /// <summary>
        /// Exclude threaded comments from being copied. 
        /// </summary>
        ExcludeThreadedComments = 0x10,
        /// <summary>
        /// Exclude hyperlinks from being copied. 
        /// </summary>
        ExcludeHyperLinks = 0x20,
        /// <summary>
        /// Exclude merged cells from being copied. 
        /// </summary>
        ExcludeMergedCells = 0x40,
        /// <summary>
        /// Exclude data validations from being copied. 
        /// </summary>
        ExcludeDataValidations = 0x80,
        /// <summary>
        /// Exclude conditional formatting from being copied. 
        /// </summary>
        ExcludeConditionalFormatting = 0x100,
        /// <summary>
        /// Transpose the copied data
        /// </summary>
        Transpose = 0x200,
        /// <summary>
        /// Exclude drawings from being copied
        /// </summary>
        ExcludeDrawings = 0x400,
        /// <summary>
        /// Exclude any table within the range. 
        /// </summary>
        ExcludeTables = 0x800,
        /// <summary>
        /// Exclude any pivot table within the range. 
        /// EPPlus will only copy pivot tables within the same workbooks.
        /// </summary>
        ExcludePivotTables = 0x1000,
        /// <summary>
        /// Exclude hidden cells in the range.
        /// </summary>
        ExcludeHiddenCells = 0x2000,
        /// <summary>
        /// Fill range with repeated data. The desination ranges rows and columns needs to be a multiple of the source's ranges rows and columns.
        /// </summary>
        Fill = 0x4000,
        /// <summary>
        /// Exclude and local cell pictures within the range.
        /// </summary>
        ExcludeLocalCellPictures = 0x8000,
        /// <summary>
        /// Exclude any web pictures (i.e. added via the IMAGE function) within the range.
        /// </summary>
        ExcludeWebPictures = 0x10000,
    }

    /// <summary>
    /// Util const for getting only all exclude flags. Without special options like Fill, Transpose etc.
    /// </summary>
    internal struct Exclude
    {
        internal const ExcelRangeCopyOptionFlags All = ExcelRangeCopyOptionFlags.ExcludeFormulas |
           ExcelRangeCopyOptionFlags.ExcludeValues |
           ExcelRangeCopyOptionFlags.ExcludeStyles |
           ExcelRangeCopyOptionFlags.ExcludeComments |
           ExcelRangeCopyOptionFlags.ExcludeThreadedComments |
           ExcelRangeCopyOptionFlags.ExcludeHyperLinks |
           ExcelRangeCopyOptionFlags.ExcludeMergedCells |
           ExcelRangeCopyOptionFlags.ExcludeDataValidations |
           ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting |
           ExcelRangeCopyOptionFlags.ExcludeTables |
           ExcelRangeCopyOptionFlags.ExcludePivotTables |
           ExcelRangeCopyOptionFlags.ExcludeLocalCellPictures |
           ExcelRangeCopyOptionFlags.ExcludeWebPictures;
    }

    /// <summary>
    /// <para>Flags to only copy certain parts of a range.</para>
    /// This provides the options of Excel's "Paste Special"
    /// </summary>
    [Flags]
    public enum ExcelRangeCopyOnly : int
    {
        //CLARIFICATRION:
        //Uses bitwise `& ~(SomeFlag)` to remove an ExcludeFlag from Exclude All.
        //Meaning only that option is copied.

        /// <summary>
        /// Paste only formulas.
        /// </summary>
        Formulas = (Exclude.All & ~ExcelRangeCopyOptionFlags.ExcludeFormulas) & ~ExcelRangeCopyOptionFlags.ExcludeValues,
        /// <summary>
        /// Paste only values
        /// </summary>
        Values = Exclude.All & ~ExcelRangeCopyOptionFlags.ExcludeValues,
        /// <summary>
        /// Paste only formatting
        /// </summary>
        Formats = (Exclude.All & ~ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting) & ~ExcelRangeCopyOptionFlags.ExcludeStyles,
        /// <summary>
        /// Paste only Comments and Threaded Comments
        /// </summary>
        Comments = (Exclude.All & ~ExcelRangeCopyOptionFlags.ExcludeComments) & ~ExcelRangeCopyOptionFlags.ExcludeThreadedComments,
        /// <summary>
        /// Paste only validation
        /// </summary>
        Validations = Exclude.All & ~ExcelRangeCopyOptionFlags.ExcludeDataValidations,

        //TODO:
        //All using source theme
        //All except borders
        //Column Widths
        //Formulas and number formats
        //Values and number formats
        //All merging conditional formats
    }
}
