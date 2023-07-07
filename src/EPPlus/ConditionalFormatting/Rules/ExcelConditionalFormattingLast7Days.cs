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
    public class ExcelConditionalFormattingLast7Days: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingLast7Days(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(eExcelConditionalFormattingRuleType.Last7Days, address, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.Last7Days;
            Formula = string.Format(
            "AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())",
            Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingLast7Days(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(eExcelConditionalFormattingRuleType.Last7Days, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingLast7Days(ExcelConditionalFormattingLast7Days copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingLast7Days(this);
        }
        #endregion
    }
}