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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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

        HashSet<object> _matchingValues;

        internal override void RemoveTempExportData()
        {
            base.RemoveTempExportData();
            _matchingValues = null;
        }

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            var cellValue = _ws.Cells[address.Address].Value;
            if (cellValue.IsNumeric())
            {
                if(_cellValueCacheUpdated == false || _matchingValues == null)
                {
                    UpdateCellValueCache();
                    //Sort once per export and keep the qualifying values. Sorting per cell is O(n^2 log n).
                    _matchingValues = GetMatchingValues();
                }

                return _matchingValues != null && _matchingValues.Contains(cellValue);
            }
            return false;
        }

        private HashSet<object> GetMatchingValues()
        {
            IEnumerable<object> values;
            switch (Type)
            {
                case eExcelConditionalFormattingRuleType.Top:
                    values = cellValueCache.Where(n => n.IsNumeric()).OrderByDescending(n => n).Take(Rank);
                    break;
                case eExcelConditionalFormattingRuleType.TopPercent:
                    values = cellValueCache.Where(n => n.IsNumeric()).OrderByDescending(n => n).Take(cellValueCache.Count * Rank / 100);
                    break;
                case eExcelConditionalFormattingRuleType.Bottom:
                    values = cellValueCache.Where(n => n.IsNumeric()).OrderBy(n => n).Take(Rank);
                    break;
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    values = cellValueCache.Where(n => n.IsNumeric()).OrderBy(n => n).Take(cellValueCache.Count * Rank / 100);
                    break;
                default:
                    return new HashSet<object>();
            }
            return new HashSet<object>(values);
        }
    }
}
