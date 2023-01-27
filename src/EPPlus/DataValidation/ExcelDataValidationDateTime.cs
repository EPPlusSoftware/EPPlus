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
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Validation for <see cref="DateTime"/>.
    /// </summary>
    public class ExcelDataValidationDateTime : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>, IExcelDataValidationDateTime
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationDateTime(string uid, string address, string worksheetName)
            : base(uid, address, worksheetName)
        {
            Formula = new ExcelDataValidationFormulaDateTime(null, Uid, worksheetName, OnFormulaChanged);
            Formula2 = new ExcelDataValidationFormulaDateTime(null, Uid, worksheetName, OnFormulaChanged);
        }

        internal ExcelDataValidationDateTime(XmlReader xr)
            : base(xr)
        {

        }

        internal override IExcelDataValidationFormulaDateTime DefineFormulaClassType(string formulaValue, string sheetName)
        {
            return new ExcelDataValidationFormulaDateTime(formulaValue, Uid, sheetName, OnFormulaChanged);
        }
        public override ExcelDataValidationType ValidationType => new ExcelDataValidationType(eDataValidationType.DateTime);
    }
}
