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
        /// <param name="worksheetName"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        internal ExcelDataValidationDateTime(string uid, string address, ExcelWorksheet ws)
            : base(uid, address, ws)
        {
            Formula = new ExcelDataValidationFormulaDateTime(null, Uid, ws.Name, OnFormulaChanged);
            Formula2 = new ExcelDataValidationFormulaDateTime(null, Uid, ws.Name, OnFormulaChanged);
        }

        /// <summary>
        /// Constructor for reading data
        /// </summary>
        /// <param name="xr">The XmlReader to read from</param>
        internal ExcelDataValidationDateTime(XmlReader xr, ExcelWorksheet ws)
            : base(xr, ws)
        {
        }

        /// <summary>
        /// Copy constructor
        /// </summary>
        /// <param name="copy"></param>
        internal ExcelDataValidationDateTime(ExcelDataValidationDateTime copy, ExcelWorksheet ws) : base(copy, ws)
        {
            Formula = copy.Formula;
            Formula2 = copy.Formula;
        }

        /// <summary>
        /// Property for determining type of validation
        /// </summary>
        public override ExcelDataValidationType ValidationType => new ExcelDataValidationType(eDataValidationType.DateTime);

        internal override IExcelDataValidationFormulaDateTime DefineFormulaClassType(string formulaValue, string sheetName)
        {
            return new ExcelDataValidationFormulaDateTime(formulaValue, Uid, sheetName, OnFormulaChanged);
        }

        internal override ExcelDataValidation GetClone()
        {
            return new ExcelDataValidationDateTime(this, _ws);
        }

        internal override ExcelDataValidation GetClone(ExcelWorksheet copy)
        {
            return new ExcelDataValidationDateTime(this, copy);
        }

        ExcelDataValidationDateTime Clone()
        {
            return (ExcelDataValidationDateTime)GetClone();
        }
    }
}
