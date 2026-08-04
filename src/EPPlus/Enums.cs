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
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Tells how cells should be shifted in a delete operation
    /// </summary>
    public enum eShiftTypeDelete
    {
        /// <summary>
        /// Cells in the range are shifted to the left
        /// </summary>
        Left,
        /// <summary>
        /// Cells in the range are shifted upwards
        /// </summary>
        Up,
        /// <summary>
        /// The range for the entire row is used in the shift operation
        /// </summary>
        EntireRow,
        /// <summary>
        /// The range for the entire column is used in the shift operation
        /// </summary>
        EntireColumn
    }
    /// <summary>
    /// Tells how cells should be shifted in a insert operation
    /// </summary>
    public enum eShiftTypeInsert
    {
        /// <summary>
        /// Cells in the range are shifted to the right
        /// </summary>
        Right,
        /// <summary>
        /// Cells in the range are shifted downwards
        /// </summary>
        Down,
        /// <summary>   
        /// The range for the entire row is used in the shift operation
        /// </summary>
        EntireRow,
        /// <summary>
        /// The range for the entire column is used in the shift operation
        /// </summary>
        EntireColumn
    }
    /// <summary>
    /// Algorithm for password hash
    /// </summary>
    internal enum eProtectedRangeAlgorithm
    {
        /// <summary>
        /// Specifies that the MD2 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        MD2,
        /// <summary>
        /// Specifies that the MD4 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        MD4,
        /// <summary>
        /// Specifies that the MD5 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        MD5,
        /// <summary>
        /// Specifies that the RIPEMD-128 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        RIPEMD128,
        /// <summary>
        /// Specifies that the RIPEMD-160 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        RIPEMD160,
        /// <summary>
        /// Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA1,
        /// <summary>
        /// Specifies that the SHA-256 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA256,
        /// <summary>
        /// Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA384,
        /// <summary>
        /// Specifies that the SHA-512 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA512,
        /// <summary>
        /// Specifies that the WHIRLPOOL algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        WHIRLPOOL
    }
    /// <summary>
    /// Maps to DotNetZips CompressionLevel enum
    /// </summary>
    public enum CompressionLevel
    {
        /// <summary>
        /// Level 0, no compression
        /// </summary>
        Level0 = 0,
        /// <summary>
        /// No compression
        /// </summary>
        None = 0,
        /// <summary>
        /// Level 1, Best speed
        /// </summary>
        Level1 = 1,
        /// <summary>
        /// 
        /// </summary>
        BestSpeed = 1,
        /// <summary>
        /// Level 2
        /// </summary>
        Level2 = 2,
        /// <summary>
        /// Level 3
        /// </summary>
        Level3 = 3,
        /// <summary>
        /// Level 4
        /// </summary>
        Level4 = 4,
        /// <summary>
        /// Level 5
        /// </summary>
        Level5 = 5,
        /// <summary>
        /// Level 6
        /// </summary>
        Level6 = 6,
        /// <summary>
        /// Default, Level 6
        /// </summary>
        Default = 6,
        /// <summary>
        /// Level 7
        /// </summary>
        Level7 = 7,
        /// <summary>
        /// Level 8
        /// </summary>
        Level8 = 8,
        /// <summary>
        /// Level 9
        /// </summary>
        BestCompression = 9,
        /// <summary>
        /// Best compression, Level 9
        /// </summary>
        Level9 = 9,
    }
    /// <summary>
    /// The position of the pane.
    /// </summary>
    public enum ePanePosition
    {
        /// <summary>
        /// Bottom Left Pane.
        /// Used when worksheet view has both vertical and horizontal splits.
        /// Also used when the worksheet is horizontally split only, specifying this is the bottom pane.
        /// </summary>
        BottomLeft,
        /// <summary>
        /// Bottom Right Pane. 
        /// This property is only used when the worksheet has both vertical and horizontal splits.
        /// </summary>
        BottomRight,
        /// <summary>
        /// Top Left Pane.
        /// Used when worksheet view has both vertical and horizontal splits.
        /// Also used when the worksheet is horizontally split only, specifying this is the top pane.
        /// </summary>
        TopLeft,
        /// <summary>
        /// Top Right Pane
        /// Used when the worksheet view has both vertical and horizontal splits.
        /// Also used when the worksheet is vertically split only, specifying this is the right pane.
        /// </summary>
        TopRight
    }
    /// <summary>
    /// The state of the pane.
    /// </summary>
    public enum ePaneState
    {
        /// <summary>
        /// Panes are frozen, but were not split being frozen.In this state, when the panes are unfrozen again, a single pane results, with no split. In this state, the split bars are not adjustable.
        /// </summary>
        Frozen,
        /// <summary>
        /// Frozen Split
        /// Panes are frozen and were split before being frozen. In this state, when the panes are unfrozen again, the split remains, but is adjustable.
        /// </summary>
        FrozenSplit,
        /// <summary>
        /// Panes are split, but not frozen.In this state, the split bars are adjustable by the user.
        /// </summary>
        Split
    }
    /// <summary>
    /// Obsolete: Specified the license EPPlus is used in versions prior to EPPlus 8.
    /// License type must be specified in order to use the library
    /// <seealso cref="ExcelPackage.LicenseContext"/>
    /// </summary>
    [Obsolete("Used in versions prior to EPPlus 8. Will be removed in coming versions.")]
    public enum LicenseContext
    {
        /// <summary>
        /// You comply with the Polyform Non Commercial License.
        /// See https://polyformproject.org/licenses/noncommercial/1.0.0/
        /// </summary>
        NonCommercial = 0,
        /// <summary>
        /// You have a commercial license purchased at https://epplussoftware.com/licenseoverview
        /// </summary>
        Commercial = 1
    }
    /// <summary>
    ///  Represents the visible state of the workbook window.
    /// </summary>
    public enum eWorkbookVisibility
    {
        /// <summary>
        /// The workbook window is hidden, but can be shown by the user via the user interface.
        /// </summary>
        Hidden,
        /// <summary>
        /// The workbook window is hidden and cannot be shown in the user interface. This state is only available programmatically.
        /// </summary>
        VeryHidden,
        /// <summary>
        /// The workbook window is visible.
        /// </summary>
        Visible
    }
    /// <summary>
    /// Specifies the document format to use when saving a package.
    /// </summary>
    public enum eDocumentFormat
    {
        /// <summary>
        /// A standard Excel workbook (.xlsx, or .xlsm if the workbook contains a VBA project).
        /// This is the default format.
        /// </summary>
        Workbook,
        /// <summary>
        /// An Excel template (.xltx, or .xltm if the workbook contains a VBA project).
        /// When opened in Excel, a template creates a new workbook based on its contents 
        /// rather than opening the template file itself for editing.
        /// </summary>
        Template
    }
}
