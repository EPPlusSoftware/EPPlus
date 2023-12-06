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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingUniqueValues : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingUniqueValues
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingUniqueValues(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.UniqueValues, address, priority, worksheet)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingUniqueValues(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.UniqueValues, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingUniqueValues(ExcelConditionalFormattingUniqueValues copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
            Rank = copy.Rank;
        }

        HashSet<object> uniqueHash = new HashSet<object>();

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            List<object> cellValues = new List<object>();
            foreach(var cell in Address.GetAllAddresses()) 
            {
                for (int i = 1; i <= cell.Rows; i++)
                {
                    for (int j = 1; j <= cell.Columns; j++)
                    {
                        cellValues.Add(_ws.Cells[cell._fromRow + i - 1, cell._fromCol + j - 1].Value);
                        //uniqueDict.Add(_ws.Cells[cell._fromRow + i - 1, cell._fromCol + j - 1].Value, $"{cell._fromRow + i - 1},{cell._fromCol + j - 1}");
                    }
                }
            }

            var uniques = cellValues.GroupBy(i => i)
                .Where(g => g.Count() == 1)
                .Select(g => g.First());
            
            if(uniques.Contains(_ws.Cells[address.Address].Value))
            {
                return true;
            }

            return false;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingUniqueValues(this, newWs);
        }

        #endregion Constructors
    }
}
