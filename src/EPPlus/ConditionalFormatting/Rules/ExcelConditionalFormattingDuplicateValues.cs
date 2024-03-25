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
using System.ComponentModel;
using System.Linq;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingDuplicateValues : CachingCF,
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

        //HashSet<string> duplicates = new HashSet<string>();
        IEnumerable<string> duplicates;

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if(cellValueCache.Count == 0)
            {
                UpdateCellValueCache(true);
                duplicates = cellValueCache.GroupBy(value => value.ToString().ToUpper())
                                           .Where(key => key.Count() > 1)
                                           .Select(group => group.Key);
            }

            if(_ws.Cells[address.Address].Value != null)
            {
                var cellVal = _ws.Cells[address.Address].Value.ToString();
                return duplicates.Contains(_ws.Cells[address.Address].Value.ToString().ToUpper());
            }
            return false;
        }

        internal override void RemoveTempExportData()
        {
            base.RemoveTempExportData();
            duplicates = null;
        }

        internal ExcelConditionalFormattingDuplicateValues(ExcelConditionalFormattingDuplicateValues copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingDuplicateValues(this, newWs);
        }
    }
}
