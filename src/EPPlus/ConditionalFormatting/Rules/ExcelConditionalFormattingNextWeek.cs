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
    public class ExcelConditionalFormattingNextWeek: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingNextWeek(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(eExcelConditionalFormattingRuleType.NextWeek, address, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.NextWeek;
            Formula = string.Format(
              "AND(ROUNDDOWN({0},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({0},0)-TODAY()<(15-WEEKDAY(TODAY())))",
              Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingNextWeek(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(eExcelConditionalFormattingRuleType.NextWeek, address, ws, xr)
        {
        }
        #endregion
    }
}