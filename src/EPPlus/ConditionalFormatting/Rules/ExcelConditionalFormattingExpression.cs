using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class ExcelConditionalFormattingExpression : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingExpression
    {
        internal ExcelConditionalFormattingExpression(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.Expression,
                address,
                priority,
                worksheet
                )
        {
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingExpression(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.Expression,
                address,
                worksheet,
                xr)
        {
        }

        internal ExcelConditionalFormattingExpression(ExcelConditionalFormattingExpression copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingExpression(this);
        }
    }
}
