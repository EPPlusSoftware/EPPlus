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
using OfficeOpenXml.Constants;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Utils;
using System;
using System.Xml;
using static OfficeOpenXml.ExcelWorksheet;

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
        /// <param name="workSheetName"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        internal ExcelDataValidationWithFormula(string uid, string address, string workSheetName)
            : base(uid, address)
        {
            _workSheetName = workSheetName;
        }

        /// <summary>
        /// Constructor for reading data
        /// </summary>
        /// <param name="xr">The XmlReader to read from</param>
        internal ExcelDataValidationWithFormula(XmlReader xr)
            : base(xr)
        {
        }

        /// <summary>
        /// Copy Constructor
        /// </summary>
        /// <param name="copy"></param>
        internal ExcelDataValidationWithFormula(ExcelDataValidation copy)
            : base(copy)
        {
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);
            Formula = ReadFormula(xr, "formula1");
        }

        internal T ReadFormula(XmlReader xr, string formulaIdentifier)
        {
            xr.ReadUntil(formulaIdentifier, "dataValidation", "extLst");

            if (xr.LocalName != formulaIdentifier && formulaIdentifier != "formula2")
                throw new NullReferenceException($"Cannot find DataValidation formula for {Uid}. " +
                    $"Missing node name: {formulaIdentifier}");

            bool isExt = xr.NamespaceURI == ExcelPackage.schemaMainX14;

            if (InternalValidationType == InternalValidationType.ExtLst || isExt)
                xr.Read();

            return DefineFormulaClassType(xr.ReadString(), _workSheetName);
        }

        abstract internal T DefineFormulaClassType(string formulaValue, string worksheetName);

        /// <summary>
        /// Formula - Either a {T} value (except for custom validation) or a spreadsheet formula
        /// </summary>
        public T Formula
        {
            get;
            protected set;
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

            var formula = Formula as ExcelDataValidationFormula;

            if (ValidationType.Type == eDataValidationType.Any)
            {
                return;
            }

            if (ValidationType.Type == eDataValidationType.List)
            {
                var formulaList = (ExcelDataValidationFormulaList)formula;
                if (string.IsNullOrEmpty(formulaList.ExcelFormula) && formulaList.Values.Count <= 0)
                {
                    throw new InvalidOperationException($"Formula in Datavalidation of type: {eDataValidationType.List} with Uid: {Uid} must have a Value or ExcelFormula");
                }
                return;
            }

            if ((formula.HasValue == false) 
                && string.IsNullOrEmpty(formula.ExcelFormula) 
                && !(AllowBlank ?? false))
            {
                throw new InvalidOperationException($"Formula in Datavalidation of type: {ValidationType.Type} with Uid: {Uid} must have a Value or ExcelFormula");
            }
        }
    }
}
