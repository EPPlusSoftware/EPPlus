﻿/*************************************************************************************************
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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingEqual : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingEqual
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingEqual(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.Equal, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.Equal;
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingEqual(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.Equal, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.Equal;
        }
        internal ExcelConditionalFormattingEqual(ExcelConditionalFormattingEqual copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
        }

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if (Address.Collide(address) != ExcelAddressBase.eAddressCollition.No)
            {
                if (_ws.Cells[address.Start.Address].Value != null && string.IsNullOrEmpty(Formula) == false)
                {
                    calculatedFormula1 = string.Format(_ws.Workbook.FormulaParserManager.Parse(GetCellFormula(address), address.FullAddress, false).ToString(), CultureInfo.InvariantCulture);
                    var str = string.Format(_ws.Cells[address.Start.Address].Value.ToString(), CultureInfo.InvariantCulture);
                    if (str == calculatedFormula1)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingEqual(this, newWs);
        }


        #endregion Constructors
    }
}
