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
using System.Xml;

namespace OfficeOpenXml.DataValidation.Formulas
{
    /// <summary>
    /// 
    /// </summary>
    internal class ExcelDataValidationFormulaCustom : ExcelDataValidationFormula, IExcelDataValidationFormula
    {
        public ExcelDataValidationFormulaCustom(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath, string validationUid)
            : base(namespaceManager, topNode, formulaPath, validationUid)
        {
            SchemaNodeOrder = new string[] { "formula1", "sqref" };
            var value = GetXmlNodeString(formulaPath);
            if (!string.IsNullOrEmpty(value))
            {
                ExcelFormula = value;
            }
            State = FormulaState.Formula;
        }

        internal override string GetXmlValue()
        {
            return ExcelFormula;
        }

        protected override string GetValueAsString()
        {
            return ExcelFormula;
        }

        internal override void ResetValue()
        {
            
        }
    }
}
