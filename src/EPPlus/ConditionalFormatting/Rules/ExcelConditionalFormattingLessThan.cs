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
    internal class ExcelConditionalFormattingLessThan : ExcelConditionalFormattingRule, IExcelConditionalFormattingLessThan
    {
        internal ExcelConditionalFormattingLessThan(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.LessThan, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.LessThan;
        }

        internal ExcelConditionalFormattingLessThan(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.LessThan, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.LessThan;
        }

        internal ExcelConditionalFormattingLessThan(ExcelConditionalFormattingLessThan copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingLessThan(this, newWs);
        }
    }
}
