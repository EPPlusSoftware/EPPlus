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
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Describes how to split a CSV text. Used by the ExcelRange.LoadFromText method.
    /// Base class for ExcelTextFormat and ExcelOutputTextFormat
    /// <seealso cref="ExcelTextFormat"/>
    /// <seealso cref="ExcelOutputTextFormat"/>
    /// </summary>
    public abstract class ExcelTextFormatBase : ExcelTextFileFormat
    {
        /// <summary>
        /// Creates a new instance if ExcelTextFormatBase
        /// </summary>
        internal ExcelTextFormatBase() : base()
        {
            DataTypes = null;
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
        /// Datatypes list for each column (if column is not present Unknown is assumed)
        /// </summary>
        public eDataTypes[] DataTypes { get; set; }
    }
}
