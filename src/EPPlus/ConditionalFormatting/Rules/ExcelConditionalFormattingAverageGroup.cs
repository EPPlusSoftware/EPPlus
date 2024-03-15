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
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingAverageGroup : ExcelConditionalFormattingRule,
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

        internal ExcelConditionalFormattingAverageGroup(ExcelConditionalFormattingAverageGroup copy, ExcelWorksheet ws = null) : base(copy, ws) 
        {
        }

        //TODO: This needs to be cleared if this CF ever changes.
        double? average = null;

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if(address != null && _ws.Cells[address.Address].Value != null)
            {
                if (_ws.Cells[address.Address].Value.IsNumeric())
                {
                    if (average == null) 
                    {
                        var avgFormula = $"AVERAGE({Address})";

                        var avgResult = _ws.Workbook.FormulaParserManager.Parse(avgFormula, address.FullAddress, false).ToString();

                        average = double.TryParse(avgResult, out double avgDouble) ? avgDouble : null;
                    }

                    if (average != null)
                    {
                        var addressValue = Convert.ToDouble(_ws.Cells[address.Address].Value);

                        switch (Type)
                        {
                            case eExcelConditionalFormattingRuleType.AboveAverage:
                                if (addressValue > average)
                                {
                                    return true;
                                }
                                break;
                            case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                                if (addressValue >= average)
                                {
                                    return true;
                                }
                                break;
                            case eExcelConditionalFormattingRuleType.BelowAverage:
                                if (addressValue < average)
                                {
                                    return true;
                                }
                                break;
                            case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                                if (addressValue <= average)
                                {
                                    return true;
                                }
                                break;
                        }
                    }
                }
            }
            return false;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet ws = null)
        {
            return new ExcelConditionalFormattingAverageGroup(this, ws);
        }
    }
}
