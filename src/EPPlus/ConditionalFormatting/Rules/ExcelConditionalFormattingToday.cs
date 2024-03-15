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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// ExcelConditionalFormattingLast7Days
    /// </summary>
    internal class ExcelConditionalFormattingToday: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingToday(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(eExcelConditionalFormattingRuleType.Today, address, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.Today;
            _baseFormula = "FLOOR({0},1)=TODAY()";
            Formula = string.Format(_baseFormula, Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingToday(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(eExcelConditionalFormattingRuleType.Today, address, ws, xr)
        {
        }
        #endregion
    }
}