﻿using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Represents a data validation mapped to the extLst element in the worksheet xml.
    /// </summary>
    public class ExcelDataValidationExtList : ExcelDataValidationWithFormula<IExcelDataValidationFormulaList>, IExcelDataValidationList
    {
        private const string _formula1ExtList = "x14:formula1/xm:f";
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationExtList(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType)
            : base(worksheet, uid, address, validationType, null, InternalValidationType.ExtLst)
        {
            Formula = new ExcelDataValidationFormulaList(NameSpaceManager, TopNode, _formula1ExtList, uid);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelDataValidationExtList(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, uid, address, validationType, itemElementNode, InternalValidationType.ExtLst)
        {
            Formula = new ExcelDataValidationFormulaList(NameSpaceManager, TopNode, _formula1ExtList, uid);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        /// <param name="namespaceManager">Namespace manager, for test purposes</param>
        internal ExcelDataValidationExtList(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, uid, address, validationType, itemElementNode, namespaceManager)
        {
            Formula = new ExcelDataValidationFormulaList(NameSpaceManager, TopNode, _formula1ExtList, uid);
        }

        private readonly string _showDropDownPath = "@showDropDown";

        internal override void RegisterFormulaListener(DataValidationFormulaListener listener)
        {
            ((ExcelDataValidationFormulaList)Formula).RegisterFormulaListener(listener);
        }

        /// <summary>
        /// True if an in-cell dropdown should be hidden.
        /// </summary>
        /// <remarks>
        /// This property corresponds to the showDropDown attribute of a data validation in Office Open Xml. Strangely enough this
        /// attributes hides the in-cell dropdown if it is true and shows the dropdown if it is not present or false. We have checked
        /// this in both Ms Excel and Google sheets and it seems like this is how it is implemented in both applications. Hence why we have
        /// renamed this property to HideDropDown since that better corresponds to the functionality.
        /// </remarks>
        public bool? HideDropDown
        {
            get
            {
                return GetXmlNodeBoolNullable(_showDropDownPath);
            }
            set
            {
                CheckIfStale();
                SetNullableBoolValue(_showDropDownPath, value);
            }
        }
    }
}
