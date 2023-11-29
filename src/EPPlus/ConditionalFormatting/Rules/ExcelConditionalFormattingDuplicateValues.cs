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
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingDuplicateValues : ExcelConditionalFormattingRule,
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

        private bool createSet = true;
        HashSet<string> values = null;

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if(createSet)
            {
                createSet = false;
                values = new HashSet<string>();
                var list = new List<string>();

                var range = new ExcelRange(_ws, _address.Address);
                foreach(var cell in range)
                {
                    if(cell.Value != null)
                    {
                        if (string.IsNullOrEmpty(cell.Value.ToString()) == false)
                        {
                            if (list.Contains(cell.Value.ToString()))
                            {
                                values.Add(cell.Value.ToString());
                            }
                            else
                            {
                                list.Add(cell.Value.ToString());
                            }
                        }
                    }
                }
                list.Clear();
            }

            if(_ws.Cells[address.Address].Value != null)
            {
                return values.Contains(_ws.Cells[address.Address].Value.ToString());
            }
            return false;
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
