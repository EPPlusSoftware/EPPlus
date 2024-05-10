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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Table;
namespace OfficeOpenXml
{
    /// <summary>
    /// Describes how to split a text file. Used by the ExcelRange.LoadFromText method.
    /// Base class for ExcelTextFormatBase, ExcelTextFormatFixedWidthBase
    /// <seealso cref="ExcelTextFormatBase"/>
    /// <seealso cref="ExcelTextFormatFixedWidthBase"/>
    /// </summary>
    public abstract class ExcelTextFileFormat
    {
        /// <summary>
        /// 
        /// </summary>
        public ExcelTextFileFormat() 
        {
        }
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
        /// <summary>
        /// Will be called for each row. Should return true if the row should be used in the export/import, otherwise false
        /// </summary>
        public Func<string, bool> ShouldUseRow { get; set; } = null;
        /// <summary>
        /// Set if data should be transposed
        /// </summary>
        public bool Transpose { get; set; } = false;
        /// <summary>
        /// If not null, create a table from the import with this style.
        /// </summary>
        public TableStyles? TableStyle { get; set; } = null;
        /// <summary>
        /// The first row used contains the headers. Will be used if the import has a <see cref="TableStyle">TableStyle</see> set.
        /// </summary>
        public bool FirstRowIsHeader { get; set; } = false;
    }
}
