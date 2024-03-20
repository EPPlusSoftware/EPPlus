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

namespace OfficeOpenXml
{
    /// <summary>
    /// Describes how to split a text file. Used by the ExcelRange.LoadFromText method.
    /// Base class for ExcelTextFormatBase, ExcelTextFormatFixedWidthBase
    /// <seealso cref="ExcelTextFormatBase"/>
    /// <seealso cref="ExcelTextFormatFixedWidthBase"/>
    /// </summary>
    public abstract class ExcelAbstractTextFormat
    {
        /// <summary>
        /// 
        /// </summary>
        public ExcelAbstractTextFormat() 
        {
            DataTypes = null;
            UseColumns = null;
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
        /// Datatypes list for each column (if column is not present Unknown is assumed)
        /// </summary>
        public eDataTypes[] DataTypes { get; set; }
        /// <summary>
        /// Specify which columns to read. If not set, all columns are read.
        /// </summary>
        public bool[] UseColumns { get; set; }
        /// <summary>
        /// Will be called for each row. Should return true if the row should be used in the export/import, otherwise false
        /// </summary>
        public Func<string, bool> ShouldUseRow { get; set; } = null;
    }
}
