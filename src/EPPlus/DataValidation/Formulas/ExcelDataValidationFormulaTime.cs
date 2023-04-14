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
 *************************************************************************************************/
using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Globalization;

namespace OfficeOpenXml.DataValidation.Formulas
{
    internal class ExcelDataValidationFormulaTime : ExcelDataValidationFormulaValue<ExcelTime>, IExcelDataValidationFormulaTime
    {
        public ExcelDataValidationFormulaTime(string formula, string validationUid, string sheetName, Action<OnFormulaChangedEventArgs> extHandler)
            : base(validationUid, sheetName, extHandler)
        {
            var value = formula;
            if (!string.IsNullOrEmpty(value))
            {
                decimal time = default;
                if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out time))
                {
                    Value = new ExcelTime(time);
                }
                else
                {
                    Value = new ExcelTime();
                    ExcelFormula = value;
                }
            }
            else
            {
                Value = new ExcelTime();
            }
            Value.TimeChanged += new EventHandler(Value_TimeChanged);
        }

        void Value_TimeChanged(object sender, EventArgs e)
        {
            _formula = Value.ToExcelString();
        }

        protected override string GetValueAsString()
        {
            if (State == FormulaState.Value)
            {
                return Value.ToExcelString();
            }
            return string.Empty;
        }

        internal override void ResetValue()
        {
            Value = new ExcelTime();
        }
    }
}
