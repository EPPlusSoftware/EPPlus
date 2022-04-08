using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Custom data validation for the x14 xml section
    /// </summary>
    public class ExcelDataValidationExtCustom : ExcelDataValidationWithFormula<IExcelDataValidationFormula>, IExcelDataValidationCustom
    {
        private const string _formula1ExtList = "x14:formula1/xm:f";
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationExtCustom(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType)
            : base(worksheet, uid, address, validationType, null, InternalValidationType.ExtLst)
        {
            Formula = new ExcelDataValidationFormulaCustom(NameSpaceManager, TopNode, _formula1ExtList, uid);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelDataValidationExtCustom(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, uid, address, validationType, itemElementNode, InternalValidationType.ExtLst)
        {
            Formula = new ExcelDataValidationFormulaCustom(NameSpaceManager, TopNode, _formula1ExtList, uid);
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
        internal ExcelDataValidationExtCustom(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, uid, address, validationType, itemElementNode, namespaceManager)
        {
            Formula = new ExcelDataValidationFormulaCustom(NameSpaceManager, TopNode, _formula1ExtList, uid);
        }

        internal override void RegisterFormulaListener(DataValidationFormulaListener listener)
        {
            ((ExcelDataValidationFormulaCustom)Formula).RegisterFormulaListener(listener);
        }
    }
}
