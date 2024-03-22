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
        }


    }
}
