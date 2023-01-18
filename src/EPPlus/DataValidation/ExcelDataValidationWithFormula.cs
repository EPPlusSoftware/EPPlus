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
using System;
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

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationWithFormula(string uid, string address)
            : base(uid, address)
        {

        }

        internal ExcelDataValidationWithFormula(XmlReader xr)
            : base(xr)
        {

        }

        internal override void LoadSpecifics(XmlReader xr)
        {
            base.LoadSpecifics(xr);
            Formula = ReadFormula(xr, "formula1");
        }

        internal T ReadFormula(XmlReader xr, string formulaIdentifier)
        {
            //xr.ReadUntil(3, formulaIdentifier, "dataValidation", "extLst");

            //if (xr.LocalName != formulaIdentifier)
            //    throw new NullReferenceException("CANNOT FIND FORMULA");

            XmlNodeType type;
            string internalFormula = null;
            do
            {
                xr.Read();
                type = xr.NodeType;
                string name = xr.Name;
                string localName = xr.LocalName;

                if (type == XmlNodeType.Element)
                    if (xr.LocalName == "formula1" || xr.LocalName == "formula2")
                    {
                        string temp = xr.ReadString();
                        if (temp == "")
                        {
                            xr.Read();
                            temp = xr.ReadString();
                        }

                        internalFormula = temp;
                    }
                    else
                        throw new NullReferenceException("CANNOT FIND FORMULA");

            } while (type != XmlNodeType.Element);

            return LoadFormula(internalFormula);
        }

        abstract internal T LoadFormula(string formulaValue);

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
