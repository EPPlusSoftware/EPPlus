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
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// A validation containing a formula
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class ExcelDataValidationWithFormula<T> : ExcelDataValidation
        where T : IExcelDataValidationFormula
    {
        protected string _workSheetName;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationWithFormula(string uid, string address, string workSheetName)
            : base(uid, address)
        {
            _workSheetName = workSheetName;
        }

        internal ExcelDataValidationWithFormula(XmlReader xr)
            : base(xr)
        {

        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);
            Formula = ReadFormula(xr, "formula1");
        }

        internal protected void checkIfExtLst(string address)
        {
            if (RefersToOtherWorksheet(Formula.ExcelFormula))
            {
                InternalValidationType = InternalValidationType.ExtLst;
            }
            else
            {
                InternalValidationType = InternalValidationType.DataValidation;
            }
        }

        internal T ReadFormula(XmlReader xr, string formulaIdentifier)
        {
            xr.ReadUntil(formulaIdentifier, "dataValidation", "extLst");

            if (xr.LocalName != formulaIdentifier)
                throw new NullReferenceException("CANNOT FIND FORMULA");

            if (InternalValidationType == InternalValidationType.ExtLst)
                xr.Read();

            return DefineFormulaClassType(xr.ReadString(), _workSheetName);
        }

        abstract internal T DefineFormulaClassType(string formulaValue, string worksheetName);



        private T _internalFormula;

        /// <summary>
        /// Formula - Either a {T} value (except for custom validation) or a spreadsheet formula
        /// </summary>
        public T Formula
        {
            get { return _internalFormula; }

            protected set
            {
                _internalFormula = value;
                checkIfExtLst(_internalFormula.ExcelFormula);
            }
        }

        private bool RefersToOtherWorksheet(string address)
        {
            if (!string.IsNullOrEmpty(address) && ExcelCellBase.IsValidAddress(address))
            {
                var adr = new ExcelAddress(address);
                return !string.IsNullOrEmpty(adr.WorkSheetName) && adr.WorkSheetName != _workSheetName;
            }
            else if (!string.IsNullOrEmpty(address))
            {
                var tokens = SourceCodeTokenizer.Default.Tokenize(address, _workSheetName);
                if (!tokens.Any()) return false;
                var addressTokens = tokens.Where(x => x.TokenTypeIsSet(TokenType.ExcelAddress));
                foreach (var token in addressTokens)
                {
                    var adr = new ExcelAddress(token.Value);
                    if (!string.IsNullOrEmpty(adr.WorkSheetName) && adr.WorkSheetName != _workSheetName)
                        return true;
                }

            }
            return false;
        }

        /// <summary>
        /// Validates the configuration of the validation.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Will be thrown if invalid configuration of the validation. Details will be in the message of the exception.
        /// </exception>
        public override void Validate()
        {
            base.Validate();
            if (ValidationType.Type != eDataValidationType.List
                && ValidationType.Type != eDataValidationType.Custom
                && (Operator == ExcelDataValidationOperator.between || Operator == ExcelDataValidationOperator.notBetween))
            {
                if (string.IsNullOrEmpty(IFormula2))
                {
                    throw new InvalidOperationException("Validation of " + Address.Address + " failed: Formula2 must be set if operator is 'between' or 'notBetween'");
                }
            }
        }
    }
}
