using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingAverageGroup : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingAverageGroup
    {
        internal ExcelConditionalFormattingAverageGroup(
         eExcelConditionalFormattingRuleType type,
         ExcelAddress address,
         int priority,
         ExcelWorksheet worksheet)
         : base(type, address, priority, worksheet)
        {
        }

        internal ExcelConditionalFormattingAverageGroup(
          eExcelConditionalFormattingRuleType type, ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(type, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingAverageGroup(ExcelConditionalFormattingAverageGroup copy): base(copy) 
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingAverageGroup(this);
        }
    }
}
