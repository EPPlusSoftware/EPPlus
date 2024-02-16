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
using System.Linq;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingTopBottomGroup : CachingCF,
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

        internal ExcelConditionalFormattingTopBottomGroup(ExcelConditionalFormattingTopBottomGroup copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
            Rank = copy.Rank;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingTopBottomGroup(this, newWs);
        }

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if (_ws.Cells[address.Address].Value.IsNumeric())
            {
                if(cellValueCache.Count == 0)
                {
                    UpdateCellValueCache();
                }

                switch (Type)
                {
                    case eExcelConditionalFormattingRuleType.Top:
                        var sorted = cellValueCache.OrderByDescending(n => n.IsNumeric()).Take(Rank);
                        if (sorted.Contains(_ws.Cells[address.Address].Value))
                        {
                            return true;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.TopPercent:
                        var percentDescending = cellValueCache.OrderByDescending(n => n.IsNumeric()).Take(cellValueCache.Count * Rank / 100);
                        if (percentDescending.Contains(_ws.Cells[address.Address].Value))
                        {
                            return true;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.Bottom:
                        var bottomSorted = cellValueCache.OrderBy(n => n.IsNumeric()).Take(Rank);
                        if (bottomSorted.Contains(_ws.Cells[address.Address].Value))
                        {
                            return true;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.BottomPercent:
                        var percentAscending = cellValueCache.OrderBy(n => n.IsNumeric()).Take(cellValueCache.Count * Rank / 100);
                        if (percentAscending.Contains(_ws.Cells[address.Address].Value))
                        {
                            return true;
                        }
                        break;
                }
            }
            return false;
        }
    }
}
