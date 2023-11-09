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
using System;
using System.Xml;
namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Represents a data validation with two formulas
    /// </summary>
    /// <typeparam name="T">An instance implementing the <see cref="IExcelDataValidationFormula"></see></typeparam>
    public abstract class ExcelDataValidationWithFormula2<T> : ExcelDataValidationWithFormula<T>
        where T : IExcelDataValidationFormula
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        internal ExcelDataValidationWithFormula2(string uid, string address, ExcelWorksheet ws)
            : base(uid, address, ws)
        {
        }

        /// <summary>
        /// Constructor for reading data
        /// </summary>
        /// <param name="xr">The XmlReader to read from</param>
        /// <param name="ws">The worksheet</param>
        internal ExcelDataValidationWithFormula2(XmlReader xr, ExcelWorksheet ws)
            : base(xr, ws)
        {
        }

        /// <summary>
        /// Copy Constructor
        /// </summary>
        /// <param name="copy"></param>
        /// <param name="ws">The worksheet</param>
        internal ExcelDataValidationWithFormula2(ExcelDataValidation copy, ExcelWorksheet ws)
            : base(copy, ws)
        {
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);

            if (Operator == ExcelDataValidationOperator.between || Operator == ExcelDataValidationOperator.notBetween)
            {
                Formula2 = ReadFormula(xr, "formula2");
            }
            else
            {
                //Ensure Formula2 is not null for writing and if operator type changes.
                Formula2 = DefineFormulaClassType("", _workSheetName);
            }
        }

        /// <summary>
        /// Formula - Either a {T} value or a spreadsheet formula
        /// </summary>
        public T Formula2
        {
            get;
            protected set;
        }

        //internal override string GetXmlValue()
        //{
        //    base.get
        //}

        /// <summary>
        /// Validate the validation
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public override void Validate()
        {
            base.Validate();
            if (ValidationType.Type != eDataValidationType.List
                && ValidationType.Type != eDataValidationType.Custom
                && (Operator == ExcelDataValidationOperator.between || Operator == ExcelDataValidationOperator.notBetween))
            {
                if (string.IsNullOrEmpty(Formula2.ExcelFormula) &&
                    (Formula2 as ExcelDataValidationFormula).HasValue == false &&
                    !(AllowBlank ?? false))
                {
                    throw new InvalidOperationException("Validation of " + Address.Address + " failed: Formula2 must be set if operator is 'between' or 'notBetween'");
                }
            }
        }
    }
}