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
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingNotBetween : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingNotBetween
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingNotBetween(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.NotBetween,
                address,
                priority,
                worksheet
                )
        {
            Operator = eExcelConditionalFormattingOperatorType.NotBetween;
            Formula = string.Empty;
            Formula2 = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingNotBetween(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.NotBetween,
                address,
                worksheet,
                xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.NotBetween;
        }

        internal ExcelConditionalFormattingNotBetween(ExcelConditionalFormattingNotBetween copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingNotBetween(this, newWs);
        }

        double? largestNumFormula = null;
        double? smallestNumFormula = null;

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            var cellValue = _ws.Cells[address.Address].Value;
            if (cellValue != null)
            {
                var str = cellValue.ToString();

                calculatedFormula1 = string.Format(_ws.Workbook.FormulaParserManager.Parse(Formula, address.FullAddress, false).ToString(), CultureInfo.InvariantCulture);
                calculatedFormula2 = string.Format(_ws.Workbook.FormulaParserManager.Parse(Formula2, address.FullAddress, false).ToString(), CultureInfo.InvariantCulture);

                var Formula1IsNum = double.TryParse(calculatedFormula1, NumberStyles.None, CultureInfo.InvariantCulture, out double num1);
                var Formula2IsNum = double.TryParse(calculatedFormula2, NumberStyles.None, CultureInfo.InvariantCulture, out double num2);
                var cellValueIsNum = double.TryParse(str, NumberStyles.None, CultureInfo.InvariantCulture, out double numCellValue);

                if (Formula1IsNum && Formula2IsNum)
                {
                    if (cellValueIsNum)
                    {
                        largestNumFormula = num1 > num2 ? num1 : num2;
                        smallestNumFormula = num1 == largestNumFormula ? num2 : num1;

                        if (smallestNumFormula <= numCellValue && numCellValue <= largestNumFormula)
                        {
                            return false;
                        }
                    }
                }
                else if (Formula1IsNum == false && Formula2IsNum == false)
                {
                    var lesserOrGreater = string.Compare(calculatedFormula1, calculatedFormula2, true);

                    var lesserStr = lesserOrGreater >= 0 ? calculatedFormula2 : calculatedFormula1;
                    var greaterStr = lesserStr == calculatedFormula1 ? calculatedFormula2 : calculatedFormula1;

                    var result1 = string.Compare(str, lesserStr, true);
                    if (result1 >= 0)
                    {
                        var result2 = string.Compare(str, greaterStr, true);

                        if (result2 <= 0)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    //If we're here one formula is string another value
                    //In excel if one formula is string and another numeric then:
                    //All numbers higher or equal than numeric value is considered applicable.
                    //While all strings compared less or equal is considered applicable.
                    double compareNum;
                    string compareString;

                    //Excel never applies cf if error is one of the formulas
                    if (calculatedFormula1[0] == '#')
                    {
                        if (ExcelErrorValue.IsErrorValue(calculatedFormula1))
                        {
                            return false;
                        }
                    }
                    //Excel never applies cf if error is one of the formulas
                    if (calculatedFormula2[0] == '#')
                    {
                        if (ExcelErrorValue.IsErrorValue(calculatedFormula2))
                        {
                            return false;
                        }
                    }

                    if (Formula1IsNum)
                    {
                        compareNum = num1;
                        compareString = calculatedFormula2;
                    }
                    else
                    {
                        compareNum = num2;
                        compareString = calculatedFormula1;
                    }

                    if (cellValueIsNum)
                    {
                        if (numCellValue >= compareNum)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        if (string.Compare(str, compareString, true) <= 0)
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        #endregion Constructors

        /****************************************************************************************/
    }
}
