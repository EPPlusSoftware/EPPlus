/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Data.Connection
{
    public class ExcelTextProperties
    {
        /// <summary>
        /// Indicates whether to prompt the user. Defaults to true.
        /// </summary>
        public bool Prompt { get; set; } = true;

        /// <summary>
        /// Specifies the file type. Defaults to "Win".
        /// </summary>
        public eConnectionTextFileType FileType { get; set; } = eConnectionTextFileType.Win;

        /// <summary>
        /// The character set to use.
        /// </summary>
        public string CharacterSet { get; set; }

        /// <summary>
        /// The first row to read. Defaults to 1.
        /// </summary>
        public uint FirstRow { get; set; } = 1;

        /// <summary>
        /// The source file path. Defaults to an empty string.
        /// </summary>
        public string SourceFile { get; set; } = string.Empty;

        /// <summary>
        /// Indicates whether the file is delimited. Defaults to true.
        /// </summary>
        public bool Delimited { get; set; } = true;

        /// <summary>
        /// The decimal separator. Defaults to ".".
        /// </summary>
        public string Decimal { get; set; } = ".";

        /// <summary>
        /// The thousands separator. Defaults to ",".
        /// </summary>
        public string Thousands { get; set; } = ",";

        /// <summary>
        /// Indicates if tab is a delimiter. Defaults to true.
        /// </summary>
        public bool Tab { get; set; } = true;

        /// <summary>
        /// Indicates if space is a delimiter. Defaults to false.
        /// </summary>
        public bool Space { get; set; } = false;

        /// <summary>
        /// Indicates if comma is a delimiter. Defaults to false.
        /// </summary>
        public bool Comma { get; set; } = false;

        /// <summary>
        /// Indicates if semicolon is a delimiter. Defaults to false.
        /// </summary>
        public bool Semicolon { get; set; } = false;

        /// <summary>
        /// Indicates if consecutive delimiters are treated as one. Defaults to false.
        /// </summary>
        public bool Consecutive { get; set; } = false;

        /// <summary>
        /// The text qualifier used. Defaults to double quote.
        /// </summary>
        public eConnectionTextQualifier Qualifier { get; set; } = eConnectionTextQualifier.DoubleQuote;

        /// <summary>
        /// The custom delimiter. Optional.
        /// </summary>
        public string Delimiter { get; set; }
        public List<ExcelConnectionTextField> Fields { get;  } = new List<ExcelConnectionTextField>();

    }
}