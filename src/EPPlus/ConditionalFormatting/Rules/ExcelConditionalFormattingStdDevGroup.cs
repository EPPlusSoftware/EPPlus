using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingStdDevGroup
      : ExcelConditionalFormattingRule,
      IExcelConditionalFormattingStdDevGroup
    {
        internal ExcelConditionalFormattingStdDevGroup(
         eExcelConditionalFormattingRuleType type,
         ExcelAddress address,
         int priority,
         ExcelWorksheet worksheet)
         : base(type, address, priority, worksheet)
        {
            StdDev = 1;
        }

        internal ExcelConditionalFormattingStdDevGroup(
          eExcelConditionalFormattingRuleType type, ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(type, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingStdDevGroup(ExcelConditionalFormattingStdDevGroup copy) : base(copy)
        {
            StdDev = copy.StdDev;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingStdDevGroup(this);
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);
            if(string.IsNullOrEmpty(xr.GetAttribute("stdDev")))
            {
                throw new InvalidOperationException($"Could not read stdDev value of ConditionalFormatting {this} of type: {Type} at adress {Address}. " +
                                                    $"XML corrupted or reading faulty");
            }
            StdDev = UInt16.Parse(xr.GetAttribute("stdDev"));
        }
    }
}
