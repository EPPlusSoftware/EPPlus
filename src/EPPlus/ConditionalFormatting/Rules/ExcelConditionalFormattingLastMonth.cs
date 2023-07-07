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
    public class ExcelConditionalFormattingLastMonth: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingLastMonth(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(eExcelConditionalFormattingRuleType.LastMonth, address, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.LastMonth;
            Formula = string.Format(
              "AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))",
              Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingLastMonth(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(eExcelConditionalFormattingRuleType.LastMonth, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingLastMonth(ExcelConditionalFormattingLastMonth copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingLastMonth(this);
        }
        #endregion
    }
}