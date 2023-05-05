using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.ConditionalFormatting.Rules;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingTopBottomGroup : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTopBottomGroup
    {
        internal ExcelConditionalFormattingTopBottomGroup(
         eExcelConditionalFormattingRuleType type,
         ExcelAddress address,
         int priority,
         ExcelWorksheet worksheet)
         : base(type, address, priority, worksheet)
        {
            Rank = 10;  // First 10 values
        }

        internal ExcelConditionalFormattingTopBottomGroup(
          eExcelConditionalFormattingRuleType type, ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(type, address, ws, xr)
        {
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);

            Rank = UInt16.Parse(xr.GetAttribute("rank"));
        }

        internal ExcelConditionalFormattingTopBottomGroup(ExcelConditionalFormattingTopBottomGroup copy) : base(copy)
        {
            Rank = copy.Rank;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingTopBottomGroup(this);
        }
    }
}
