/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/30/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Padding types, can be left, right or auto.
    /// </summary>
    public enum PaddingAlignmentType
    {
        /// <summary>
        /// Detects the padding type automatically. Text will be left and numbers will be right.
        /// </summary>
        Auto,
        /// <summary>
        /// Add padding to the left.
        /// </summary>
        Left,
        /// <summary>
        /// Add padding to the right.
        /// </summary>
        Right
    }

    /// <summary>
    /// Describes how to output an fixed width text file.
    /// </summary>
    public class ExcelOutputTextFormatFixedWidth : ExcelTextFormatFixedWidthBase
    {
        /// <summary>
        /// Describes how to split a fixed width text
        /// </summary>
        public ExcelOutputTextFormatFixedWidth() : base() { }

        /// <summary>
        /// A text written at the start of the file.
        /// </summary>
        public string Header { get; set; }
        /// <summary>
        /// A text written at the end of the file
        /// </summary>
        public string Footer { get; set; }
        /// <summary>
        /// Flag to exclude header for fixed width text file
        /// </summary>
        public bool ExcludeHeader { get; set; } = false;
        /// <summary>
        /// Use the cells Text property with the applied culture.
        /// This only applies to columns with no format set in the Formats collection.
        /// If SkipLinesBeginning is larger than zero, headers will still be read from the first row in the range.
        /// If a TextQualifier is set, non numeric and date columns will be wrapped with the TextQualifier
        /// </summary>
        public bool UseCellFormat { get; set; } = true;
        /// <summary>
        /// A specific .NET format for the column.
        /// Format is applied with the used culture.
        /// For a text column use $ as format
        /// </summary>        
        public string[] Formats { get; set; } = null;
        /// <summary>
        /// Decimal separator, if other than the used culture.
        /// </summary>
        public string DecimalSeparator { get; set; } = null;
        /// <summary>
        /// Thousands separator, if other than the used culture.
        /// </summary>
        public string ThousandsSeparator { get; set; } = null;
        /// <summary>
        /// What to replace Text Qualifiers inside a text, when Text Qualifiers is set.
        /// Default is two Text Qualifiers characters. For example " is replaced with "".
        /// </summary>
        public string EncodedTextQualifiers { get; set; } = null;
        /// <summary>
        /// Set this to output file with trailing minus signs.
        /// </summary>
        public bool UseTrailingMinus { get; set; } = false;
    }
}
