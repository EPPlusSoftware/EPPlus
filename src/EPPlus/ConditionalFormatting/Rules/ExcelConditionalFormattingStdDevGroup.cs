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
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Utilities;

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

        internal ExcelConditionalFormattingStdDevGroup(ExcelConditionalFormattingStdDevGroup copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
            StdDev = copy.StdDev;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingStdDevGroup(this, newWs);
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

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if (_ws.Cells[address.Address].Value == null)
            {
                return false;
            }
            if (_ws.Cells[address.Address].Value.IsNumeric() == false)
            {
                return false;
            }

            var addressValue = Convert.ToDouble(_ws.Cells[address.Address].Value);

            var stdDevFormula = $"{StdDev}*STDEV.S({Address})";
            var avgFormula = $"AVERAGE({Address})";

            var stdDevRes = _ws.Workbook.FormulaParserManager.Parse(stdDevFormula, address.FullAddress, false).ToString();
            var avgResult = _ws.Workbook.FormulaParserManager.Parse(avgFormula, address.FullAddress, false).ToString();

            var stdParsable = double.TryParse(stdDevRes, out double stdDevDouble);
            var avgParsable = double.TryParse(avgResult, out double avgDouble);

            if (!(stdParsable && avgParsable)) { return false; }

            switch (Type)
            {
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                    if(addressValue > (avgDouble + stdDevDouble))
                    {
                        return true;
                    }
                    break;
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    if (addressValue < (avgDouble + stdDevDouble))
                    {
                        return true;
                    }
                    break;

            }

            return false;
        }
    }
}
