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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation;
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.DataValidation.Formulas
{
    internal class ExcelDataValidationFormulaTime : ExcelDataValidationFormulaValue<ExcelTime>, IExcelDataValidationFormulaTime
    {
        public ExcelDataValidationFormulaTime(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath, string validationUid)
            : base(namespaceManager, topNode, formulaPath, validationUid)
        {
            var value = GetXmlNodeString(formulaPath);
            if (!string.IsNullOrEmpty(value))
            {
                decimal time = default(decimal);
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
            SetXmlNodeString(FormulaPath, Value.ToExcelString());
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
