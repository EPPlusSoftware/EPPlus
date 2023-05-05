using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingDuplicateValues : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingDuplicateValues
    {
        internal ExcelConditionalFormattingDuplicateValues(
         ExcelAddress address,
         int priority,
         ExcelWorksheet worksheet)
         : base(eExcelConditionalFormattingRuleType.DuplicateValues, address, priority, worksheet)
        {

        }

        internal ExcelConditionalFormattingDuplicateValues(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.DuplicateValues, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingDuplicateValues(ExcelConditionalFormattingDuplicateValues copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingDuplicateValues(this);
        }
    }
}
