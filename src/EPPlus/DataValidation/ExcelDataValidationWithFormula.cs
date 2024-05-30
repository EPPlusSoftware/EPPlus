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
        /// <summary>
        /// Name of worksheet this datavalidation belongs to
        /// </summary>
        internal protected string _workSheetName;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="ws">The worksheet</param>
        /// <param name="address"></param>
        internal ExcelDataValidationWithFormula(string uid, string address, ExcelWorksheet ws)
            : base(uid, address, ws)
        {
            _workSheetName = ws.Name;
        }

        /// <summary>
        /// Constructor for reading data
        /// </summary>
        /// <param name="xr">The XmlReader to read from</param>
        /// <param name="ws">The worksheet</param>
        internal ExcelDataValidationWithFormula(XmlReader xr, ExcelWorksheet ws)
            : base(xr, ws)
        {
        }

        /// <summary>
        /// Copy Constructor
        /// </summary>
        /// <param name="copy"></param>
        /// <param name="ws">The worksheet</param>

        internal ExcelDataValidationWithFormula(ExcelDataValidation copy, ExcelWorksheet ws)
            : base(copy, ws)
        {
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);
            Formula = ReadFormula(xr, "formula1");
        }

        internal T ReadFormula(XmlReader xr, string formulaIdentifier)
        {
            T retVal;

            if (xr.LocalName == "AlternateContent")
            {
                if(xr.IsEmptyElement == false)
                {
                    if (xr.ReadUntil(formulaIdentifier, "AlternateContent"))
                    {
                        xr.Read();
                        retVal = DefineFormulaClassType(xr.ReadString(), _workSheetName);
                        xr.Read();

                        xr.ReadUntil("Formula2", "dataValidation", "dataValidations");
                        return retVal;
                    }
                }
                //Read over AlternateContent end node.
                //to DataValidation or dataValidations end node
                xr.Read();
            }

            if (xr.LocalName != formulaIdentifier)
            {
                return DefineFormulaClassType(null, _workSheetName);
            }

            bool isExt = xr.NamespaceURI == ExcelPackage.schemaMainX14;

            //Old Epplus files did not read extLst properly it is possible to not be extLst before DefineFormulaClassType
            //and become extLst after. Therefore don't split the check to before and after.
            if (InternalValidationType == InternalValidationType.ExtLst || isExt)
            {
                xr.Read();

                retVal = DefineFormulaClassType(xr.ReadString(), _workSheetName);

                xr.Read();
            }
            else
            {
                retVal = DefineFormulaClassType(xr.ReadString(), _workSheetName);
            }

            xr.Read();

            return retVal;
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
        }
    }
}