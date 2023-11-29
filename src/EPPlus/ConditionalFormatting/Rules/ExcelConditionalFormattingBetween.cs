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
                // _ws.Cells[address.Address].Calculate();
                var str = cellValue.ToString();
                //var formulaResult = CalculationExtension.Calculate(_ws, str);
                //_ws.Calculate()
                var cRes = _ws.Workbook.FormulaParserManager.Parse(str);

                var formulaResult = RpnFormulaExecution.ExecuteFormula(_ws, str, new ExcelCalculationOption());

                if (formulaResult is string)
                {
                    string stuff = (string)formulaResult;
                }

                if (double.TryParse(str, NumberStyles.None, CultureInfo.InvariantCulture, out double numCellValue))
                {
                    SetNumFormulas();

                    if (largestNumFormula != null && smallestNumFormula != null)
                    {
                        if (smallestNumFormula <= numCellValue && numCellValue <= largestNumFormula)
                        {
                            return true;
                        }
                    }

                    //If we're here both numFormulas cannot have values but one can
                    //In excel if one formula is string and another numeric all numbers higher than numeric value is considered applicable.
                    if (NumFormula != null)
                    {
                        if (numCellValue >= NumFormula)
                        {
                            return true;
                        }
                    }
                    else if (NumFormula2 != null)
                    {
                        if (numCellValue >= NumFormula2)
                        {
                            return true;
                        }
                    }

                    return false;
                }

                if(NumFormula != null)
                {
                    if(NumFormula2 != null)
                    {
                        return false;
                    }
                    
                    if(string.Compare(str, Formula2, true) <= 0)
                    {
                        return true;
                    }
                }

                if (NumFormula2 != null)
                {
                    if(NumFormula != null)
                    {
                        return false;
                    }

                    if (string.Compare(str, Formula, true) <= 0)
                    {
                        return true;
                    }
                }

                var lesserOrGreater = string.Compare(Formula, Formula2, true);

                var lesserStr = lesserOrGreater >= 0 ? Formula2 : Formula;
                var greaterStr = lesserStr == Formula ? Formula2 : Formula;

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
            return false;
        }

        void SetNumFormulas()
        {
            if (largestNumFormula == null && smallestNumFormula == null)
            {
                ////Ensure getter is triggered for both.
                //var numFormula1 = NumFormula;
                //var numFormula2 = NumFormula2;

                if (NumFormula != null && NumFormula2 != null)
                {
                    largestNumFormula = NumFormula > NumFormula2 ? NumFormula : NumFormula2;
                    smallestNumFormula = NumFormula == largestNumFormula ? NumFormula2 : NumFormula;
                }
            }
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
