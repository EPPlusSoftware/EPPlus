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
    /// This class represents an List data validation.
    /// </summary>
    public class ExcelDataValidationList : ExcelDataValidationWithFormula<IExcelDataValidationFormulaList>, IExcelDataValidationList
    {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationList(string uid, string address, string worksheetName)
            : base(uid, address, worksheetName)
        {
            Formula = new ExcelDataValidationFormulaList(null, uid, worksheetName, OnFormulaChanged);
        }

        /// <summary>
        /// Constructor for reading data
        /// </summary>
        /// <param name="xr">The XmlReader to read from</param>
        internal ExcelDataValidationList(XmlReader xr)
            : base(xr)
        {
        }

        /// <summary>
        /// Copy constructor
        /// </summary>
        /// <param name="copy"></param>
        internal ExcelDataValidationList(ExcelDataValidationList copy) : base(copy)
        {
            Formula = copy.Formula;
        }

        /// <summary>
        /// Read-Only property for seeing if this dataValidation type has an operator.
        /// </summary>
        public override bool AllowsOperator { get { return false; } }

        /// <summary>
        /// Property for determining type of validation
        /// </summary>
        public override ExcelDataValidationType ValidationType => new ExcelDataValidationType(eDataValidationType.List);

        /// <summary>
        /// True if an in-cell dropdown should be hidden.
        /// </summary>
        /// <remarks>
        /// This property corresponds to the showDropDown attribute of a data validation in Office Open Xml. Strangely enough this
        /// attributes hides the in-cell dropdown if it is true and shows the dropdown if it is not present or false. We have checked
        /// this in both Ms Excel and Google sheets and it seems like this is how it is implemented in both applications. Hence why we have
        /// renamed this property to HideDropDown since that better corresponds to the functionality.
        /// </remarks>
        public bool? HideDropDown { get; set; }

        public override void Validate()
        {
            base.Validate();
        }

        internal override IExcelDataValidationFormulaList DefineFormulaClassType(string formulaValue, string sheetName)
        {
            return new ExcelDataValidationFormulaList(formulaValue, Uid, sheetName, OnFormulaChanged);
        }

        internal override void LoadXML(XmlReader xr)
        {
            base.LoadXML(xr);
            string attribute = xr.GetAttribute("showDropDown");
            if (string.IsNullOrEmpty(attribute))
            {
                HideDropDown = false;
            }
            else
            {
                HideDropDown = bool.Parse(xr.GetAttribute("showDropDown"));
            }
        }

        internal override ExcelDataValidation GetClone()
        {
            return new ExcelDataValidationList(this);
        }

        ExcelDataValidationDecimal Clone()
        {
            return (ExcelDataValidationDecimal)GetClone();
        }
    }
}
