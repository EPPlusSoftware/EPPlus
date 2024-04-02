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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingGreaterThan : ExcelConditionalFormattingRule, IExcelConditionalFormattingGreaterThan
    {

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingGreaterThan(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.GreaterThan, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThan;
            //Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingGreaterThan(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.GreaterThan, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThan;
        }

        internal ExcelConditionalFormattingGreaterThan(ExcelConditionalFormattingGreaterThan copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
        }

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            var cellValue = _ws.Cells[address.Address].Value;
            if (cellValue != null && string.IsNullOrEmpty(Formula) == false)
            {
                calculatedFormula1 = string.Format(_ws.Workbook.FormulaParserManager.Parse(GetCellFormula(address)).ToString(), CultureInfo.InvariantCulture);
                if(double.TryParse(calculatedFormula1, out double result))
                {
                    if(cellValue.IsNumeric())
                    {
                        return Convert.ToDouble(cellValue) > result;
                    }
                }
                else
                {
                    var compareResult = string.Compare(calculatedFormula1, cellValue.ToString(), true);
                    return compareResult > 0;
                }
            }

            return false;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingGreaterThan(this, newWs);
        }

        #endregion Constructors
    }
}
