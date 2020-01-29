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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;

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
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {
            Formula = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, address, validationType, itemElementNode)
        {
            Formula = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        /// <param name="namespaceManager"></param>
        internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, address, validationType, itemElementNode, namespaceManager)
        {
            Formula = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula2Path);
        }
    }
}
