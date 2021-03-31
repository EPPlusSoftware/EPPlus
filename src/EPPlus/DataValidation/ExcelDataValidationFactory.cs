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
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Factory class for ExcelDataValidation.
    /// </summary>
    internal static class ExcelDataValidationFactory
    {
        /// <summary>
        /// Creates an instance of <see cref="ExcelDataValidation"/> out of the given parameters.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="itemElementNode"></param>
        /// <param name="internalType"></param>
        /// <param name="uid"></param>
        /// <returns></returns>
        internal static ExcelDataValidation Create(ExcelDataValidationType type, ExcelWorksheet worksheet, string address, XmlNode itemElementNode, InternalValidationType internalType, string uid)
        {
            Require.Argument(type).IsNotNull("validationType");
            switch (type.Type)
            {
                case eDataValidationType.Any:
                    return new ExcelDataValidationAny(worksheet, uid, address, type, itemElementNode);
                case eDataValidationType.TextLength:
                case eDataValidationType.Whole:
                    return new ExcelDataValidationInt(worksheet, uid, address, type, itemElementNode);
                case eDataValidationType.Decimal:
                    return new ExcelDataValidationDecimal(worksheet, uid, address, type, itemElementNode);
                case eDataValidationType.List:
                    return CreateListValidation(type, worksheet, address, itemElementNode, internalType, uid);
                case eDataValidationType.DateTime:
                    return new ExcelDataValidationDateTime(worksheet, uid, address, type, itemElementNode);
                case eDataValidationType.Time:
                    return new ExcelDataValidationTime(worksheet, uid, address, type, itemElementNode);
                case eDataValidationType.Custom:
                    return CreateCustomValidation(type, worksheet, address, itemElementNode, internalType, uid);
                default:
                    throw new InvalidOperationException("Non supported validationtype: " + type.Type.ToString());
            }
        }

        internal static ExcelDataValidationWithFormula<IExcelDataValidationFormulaList> CreateListValidation(ExcelDataValidationType type, ExcelWorksheet worksheet, string address, XmlNode itemElementNode, InternalValidationType internalType, string uid)
        {
            if(internalType == InternalValidationType.DataValidation)
            {
                return new ExcelDataValidationList(worksheet, uid, address, type, itemElementNode);
            }
            // extLst
            return new ExcelDataValidationExtList(worksheet, uid, address, type, itemElementNode);
        }

        internal static ExcelDataValidationWithFormula<IExcelDataValidationFormula> CreateCustomValidation(ExcelDataValidationType type, ExcelWorksheet worksheet, string address, XmlNode itemElementNode, InternalValidationType internalType, string uid)
        {
            if (internalType == InternalValidationType.DataValidation)
            {
                return new ExcelDataValidationCustom(worksheet, uid, address, type, itemElementNode);
            }
            // extLst
            return new ExcelDataValidationExtCustom(worksheet, uid, address, type, itemElementNode);
        }

    }
}
