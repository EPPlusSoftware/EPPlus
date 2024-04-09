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
using System.Globalization;

namespace OfficeOpenXml
{
    /// <summary>
    /// Discribes a column when reading a text using the ExcelRangeBase.LoadFromText method
    /// </summary>
    public enum eDataTypes
    {
        /// <summary>
        /// Let the the import decide.
        /// </summary>
        Unknown,
        /// <summary>
        /// Always a string.
        /// </summary>
        String,
        /// <summary>
        /// Try to convert it to a number. If it fails then add it as a string.
        /// </summary>
        Number,
        /// <summary>
        /// Try to convert it to a date. If it fails then add it as a string.
        /// </summary>
        DateTime,
        /// <summary>
        /// Try to convert it to a number and divide with 100. 
        /// Removes any tailing percent sign (%). If it fails then add it as a string.
        /// </summary>
        Percent
    }
    /// <summary>
    /// Describes how to split a CSV text. Used by the ExcelRange.LoadFromText method.
    /// Base class for ExcelTextFormat and ExcelOutputTextFormat
    /// <seealso cref="ExcelTextFormat"/>
    /// <seealso cref="ExcelOutputTextFormat"/>
    /// </summary>
    public class ExcelTextFormatBase
    {
        /// <summary>
        /// Creates a new instance if ExcelTextFormatBase
        /// </summary>
        public ExcelTextFormatBase()
        {
        }
        /// <summary>
        /// Delimiter character
        /// </summary>
        public char Delimiter { get; set; } = ',';
        /// <summary>
        /// Text qualifier character. Default no TextQualifier (\0)
        /// </summary>
        public char TextQualifier { get; set; } = '\0';
        /// <summary>
        /// End of line characters. Default is CRLF
        /// </summary>
        public string EOL { get; set; } = "\r\n";
        /// <summary>
        /// Culture used when parsing. Default CultureInfo.InvariantCulture
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
        /// <summary>
        /// Number of lines skipped in the begining of the file. Default 0.
        /// </summary>
        public int SkipLinesBeginning { get; set; } = 0;
        /// <summary>
        /// Number of lines skipped at the end of the file. Default 0.
        /// </summary>
        public int SkipLinesEnd { get; set; } = 0;
        /// <summary>
        /// Only used when reading/writing files from disk using a FileInfo object. Default AscII
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.ASCII;
    }

    /// <summary>
    /// Describes how to split a CSV text. Used by the ExcelRange.LoadFromText method
    /// </summary>
    public class ExcelTextFormat : ExcelTextFormatBase
    {
        /// <summary>
        /// Describes how to split a CSV text
        /// 
        /// Default values
        /// <list>
        /// <listheader><term>Property</term><description>Value</description></listheader>
        /// <item><term>Delimiter</term><description>,</description></item>
        /// <item><term>TextQualifier</term><description>None (\0)</description></item>
        /// <item><term>EOL</term><description>CRLF</description></item>
        /// <item><term>Culture</term><description>CultureInfo.InvariantCulture</description></item>
        /// <item><term>SkipLinesBeginning</term><description>0</description></item>
        /// <item><term>SkipLinesEnd</term><description>0</description></item>
        /// <item><term>DataTypes</term><description>Column datatypes</description></item>
        /// <item><term>Encoding</term><description>Encoding.ASCII</description></item>
        /// </list>
        /// </summary>
        public ExcelTextFormat() : base()
        {
            DataTypes=null;
        }
        /// <summary>
        /// Datatypes list for each column (if column is not present Unknown is assumed)
        /// </summary>
        public eDataTypes[] DataTypes { get; set; }
    }

    /// <summary>
    /// Describes how to split a CSV text. Used by the ExcelRange.SaveFromText method
    /// </summary>
    public class ExcelOutputTextFormat : ExcelTextFormatBase
    {
        /// <summary>
        /// Describes how to split a CSV text
        /// 
        /// Default values
        /// <list>
        /// <listheader><term>Property</term><description>Value</description></listheader>
        /// <item><term>Delimiter</term><description>,</description></item>
        /// <item><term>TextQualifier</term><description>None (\0)</description></item>
        /// <item><term>EOL</term><description>CRLF</description></item>
        /// <item><term>Culture</term><description>CultureInfo.InvariantCulture</description></item>
        /// <item><term>SkipLinesBeginning</term><description>0</description></item>
        /// <item><term>SkipLinesEnd</term><description>0</description></item>
        /// <item><term>Header</term><description></description></item>
        /// <item><term>Footer</term><description></description></item>
        /// <item><term>FirstRowIsHeader</term><description>true</description></item>
        /// <item><term>Encoding</term><description>Encoding.ASCII</description></item>
        /// <item><term>UseCellFormat</term><description>true</description></item>
        /// <item><term>Formats</term><description>Formats can be .NET number format, dateformats. For text use a $. A blank formats will try to autodetect</description></item>
        /// <item><term>DecimalSeparator</term><description>From Culture(null)</description></item>
        /// <item><term>ThousandsSeparator</term><description>From Culture(null)</description></item>
        /// </list> 
        /// </summary>
        public ExcelOutputTextFormat() : base()
        {
            
        }
        /// <summary>
        /// A text written at the start of the file.
        /// </summary>
        public string Header { get; set; }
        /// <summary>
        /// A text written at the end of the file
        /// </summary>
        public string Footer { get; set; }
        /// <summary>
        /// First row of the range contains the headers.
        /// All header cells will be treated as strings.
        /// </summary>
        public bool FirstRowIsHeader { get; set; } = true;
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
        public string[] Formats { get; set; }=null;
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
        /// Set if data in worksheet is transposed.
        /// </summary>
        public bool DataIsTransposed { get; set; } = false;

    }
}
