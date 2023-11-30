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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingBetween : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingBetween
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingBetween(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.Between,
                address,
                priority,
                worksheet
                )
        {
            Operator = eExcelConditionalFormattingOperatorType.Between;
            Formula = string.Empty;
            Formula2 = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingBetween(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.Between,
                address,
                worksheet,
                xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.Between;
        }


        double? largestNumFormula = null;
        double? smallestNumFormula = null;


        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            var cellValue = _ws.Cells[address.Address].Value;
            if (cellValue != null)
            {
                var str = cellValue.ToString();

                _ws.Calculate();

                //TODO:
                //We must calculate per cell bc otherwise some formulas e.g. =ROW() will be inaccurate.
                //This is ineffectual. Perhaps we could simply flag if a formula requires recalculation or not instead.
                calculatedFormula1 = RpnFormulaExecution.ExecuteFormula(_ws, Formula, new ExcelCalculationOption()).ToString();
                calculatedFormula2 = RpnFormulaExecution.ExecuteFormula(_ws, Formula2, new ExcelCalculationOption()).ToString();

                //TODO: Should be used instead but for some reason applies the value to the cell.
                //calculatedFormula1 = _ws.Workbook.FormulaParserManager.Parse(Formula, address.FullAddress).ToString();
                //calculatedFormula2 = _ws.Workbook.FormulaParserManager.Parse(Formula2, address.FullAddress).ToString();

                //TODO: Date-Handling? Should be a double either way?
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
                            return true;
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
                            return true;
                        }
                    }
                }
                else
                {
                    //If we're here one formula is string another value
                    //In excel if one formula is string and another numeric all numbers higher than numeric value is considered applicable.
                    //While all strings compared less is considered applicable.
                    double compareNum;
                    string compareString;

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
                            return true;
                        }
                    }
                    else
                    {
                        if(string.Compare(str, compareString, true) <= 0)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        internal ExcelConditionalFormattingBetween(ExcelConditionalFormattingBetween copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingBetween(this, newWs);
        }


        #endregion Constructors

        /****************************************************************************************/
    }
}
