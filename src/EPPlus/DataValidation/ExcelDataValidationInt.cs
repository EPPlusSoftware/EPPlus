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
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Data validation for integer values.
    /// </summary>
    public class ExcelDataValidationInt : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>, Contracts.IExcelDataValidationInt
    {
        bool _isTextLength = false;

        /// <summary>
        /// Constructor for reading data
        /// </summary>
        /// <param name="xr">The XmlReader to read from</param>
        ///  <param name="isTextLength">Bool to define type of int validation</param>
        internal ExcelDataValidationInt(XmlReader xr, ExcelWorksheet ws, bool isTextLength = false)
            : base(xr, ws)
        {
            _isTextLength = isTextLength;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="isTextLength">Bool to define type of int validation</param>
        internal ExcelDataValidationInt(string uid, string address, ExcelWorksheet ws, bool isTextLength = false)
            : base(uid, address, ws)
        {
            //Initilization of forumlas so they don't cause nullref
            Formula = new ExcelDataValidationFormulaInt(null, uid, ws.Name, OnFormulaChanged);
            Formula2 = new ExcelDataValidationFormulaInt(null, uid, ws.Name, OnFormulaChanged);
            _isTextLength = isTextLength;
        }

        /// <summary>
        /// Copy constructor
        /// </summary>
        /// <param name="copy"></param>
        internal ExcelDataValidationInt(ExcelDataValidationInt copy, ExcelWorksheet ws) 
            : base(copy, ws)
        {
            Formula = copy.Formula;
            Formula2 = copy.Formula2;
        }

        /// <summary>
        /// Property for determining type of validation
        /// </summary>
        public override ExcelDataValidationType ValidationType => _isTextLength ?
            new ExcelDataValidationType(eDataValidationType.TextLength) :
            new ExcelDataValidationType(eDataValidationType.Whole);

        internal override IExcelDataValidationFormulaInt DefineFormulaClassType(string formulaValue, string worksheetName)
        {
            return new ExcelDataValidationFormulaInt(formulaValue, Uid, worksheetName, OnFormulaChanged);
        }

        internal override ExcelDataValidation GetClone()
        {
            return new ExcelDataValidationInt(this, _ws);
        }

        /// <summary>
        /// Return a deep-copy clone of validation
        /// </summary>
        /// <returns></returns>
        public ExcelDataValidationInt Clone()
        {
            return (ExcelDataValidationInt)GetClone();
        }

    }
}
