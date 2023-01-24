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
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Data validation for integer values.
    /// </summary>
    public class ExcelDataValidationInt : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>, IExcelDataValidationInt
    {

        internal ExcelDataValidationInt(XmlReader xr) : base(xr)
        {

        }

        internal override IExcelDataValidationFormulaInt DefineFormulaClassType(string formulaValue, string worksheetName)
        {
            return new ExcelDataValidationFormulaInt(formulaValue, Uid, worksheetName);
        }

        public override ExcelDataValidationType ValidationType => new ExcelDataValidationType(eDataValidationType.Whole);


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationInt(string uid, string address, string worksheetName) : base(uid, address, worksheetName)
        {
            //Initilization of forumlas so they don't cause nullref
            Formula = new ExcelDataValidationFormulaInt(null, uid, worksheetName);
            Formula2 = new ExcelDataValidationFormulaInt(null, uid, worksheetName);
        }
    }
}
