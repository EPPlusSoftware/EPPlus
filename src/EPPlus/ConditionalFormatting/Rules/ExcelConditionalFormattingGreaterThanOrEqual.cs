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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingGreaterThanOrEqual : ExcelConditionalFormattingRule, IExcelConditionalFormattingGreaterThanOrEqual
    {

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingGreaterThanOrEqual(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.GreaterThanOrEqual, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThanOrEqual;
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingGreaterThanOrEqual(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.GreaterThanOrEqual, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThanOrEqual;
        }

        internal ExcelConditionalFormattingGreaterThanOrEqual(ExcelConditionalFormattingGreaterThanOrEqual copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingGreaterThanOrEqual(this);
        }

        #endregion Constructors
    }
}
