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
