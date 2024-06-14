/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/05/2021         EPPlus Software AB           Bugfix
 *************************************************************************************************/
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    /// <summary>
    /// Range or value
    /// </summary>
    internal class RangeOrValue
    {
        /// <summary>
        /// Value
        /// </summary>
        public object Value { get; set; }
        /// <summary>
        /// Range
        /// </summary>
        public IRangeInfo Range { get; set; }
    }
}
