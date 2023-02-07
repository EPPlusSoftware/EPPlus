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
using System.Xml;

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
        ///
        internal static ExcelDataValidation Create(XmlReader xr)
        {
            string validationTypeName = xr.GetAttribute("type") == null ? "" : xr.GetAttribute("type");

            switch (validationTypeName)
            {
                case "":
                    return new ExcelDataValidationAny(xr);
                case "textLength":
                    return new ExcelDataValidationInt(xr, true);
                case "whole":
                    return new ExcelDataValidationInt(xr);
                case "decimal":
                    return new ExcelDataValidationDecimal(xr);
                case "list":
                    return new ExcelDataValidationList(xr);
                case "time":
                    return new ExcelDataValidationTime(xr);
                case "date":
                    return new ExcelDataValidationDateTime(xr);
                case "custom":
                    return new ExcelDataValidationCustom(xr);
                default:
                    throw new InvalidOperationException($"Non supported validationtype: {validationTypeName}");
            }
        }

        //TODO: Improve this by making formula copying internal or at least pressed lines via method call
        internal static ExcelDataValidation Create(ExcelDataValidation oldValidation, string address, string uid, string workSheetName)
        {
            switch (oldValidation.ValidationType.Type)
            {
                case eDataValidationType.Any:
                    return new ExcelDataValidationAny(uid, address);
                case eDataValidationType.TextLength:
                case eDataValidationType.Whole:
                    var intValidation = new ExcelDataValidationInt(uid, address, workSheetName);

                    intValidation.Formula.Value = oldValidation.As.IntegerValidation.Formula.Value;
                    intValidation.Formula.ExcelFormula = oldValidation.As.IntegerValidation.Formula.ExcelFormula;
                    intValidation.Formula2.Value = oldValidation.As.IntegerValidation.Formula2.Value;
                    intValidation.Formula2.ExcelFormula = oldValidation.As.IntegerValidation.Formula2.ExcelFormula;

                    return intValidation;
                case eDataValidationType.Decimal:

                    var decimalValidation = new ExcelDataValidationDecimal(uid, address, workSheetName);

                    decimalValidation.Formula.Value = oldValidation.As.DecimalValidation.Formula.Value;
                    decimalValidation.Formula.ExcelFormula = oldValidation.As.DecimalValidation.Formula.ExcelFormula;
                    decimalValidation.Formula2.Value = oldValidation.As.DecimalValidation.Formula2.Value;
                    decimalValidation.Formula2.ExcelFormula = oldValidation.As.DecimalValidation.Formula2.ExcelFormula;

                    return decimalValidation;
                case eDataValidationType.List:
                    var listValidation = new ExcelDataValidationList(uid, address, workSheetName);
                    for (int i = 0; i < oldValidation.As.ListValidation.Formula.Values.Count; i++)
                    {
                        listValidation.Formula.Values.Add(oldValidation.As.ListValidation.Formula.Values[i]);
                    }
                    listValidation.Formula.ExcelFormula = oldValidation.As.ListValidation.Formula.ExcelFormula;

                    return listValidation;
                case eDataValidationType.DateTime:
                    var dateTimeValidation = new ExcelDataValidationDateTime(uid, address, workSheetName);

                    dateTimeValidation.Formula.Value = oldValidation.As.DateTimeValidation.Formula.Value;
                    dateTimeValidation.Formula.ExcelFormula = oldValidation.As.DateTimeValidation.Formula.ExcelFormula;
                    dateTimeValidation.Formula2.Value = oldValidation.As.DateTimeValidation.Formula2.Value;
                    dateTimeValidation.Formula2.ExcelFormula = oldValidation.As.DateTimeValidation.Formula2.ExcelFormula;

                    return dateTimeValidation;
                case eDataValidationType.Time:
                    var timeValidation = new ExcelDataValidationTime(uid, address, workSheetName);

                    timeValidation.Formula.Value = oldValidation.As.TimeValidation.Formula.Value;
                    timeValidation.Formula.ExcelFormula = oldValidation.As.TimeValidation.Formula.ExcelFormula;
                    timeValidation.Formula2.Value = oldValidation.As.TimeValidation.Formula2.Value;
                    timeValidation.Formula2.ExcelFormula = oldValidation.As.TimeValidation.Formula2.ExcelFormula;

                    return timeValidation;
                case eDataValidationType.Custom:
                    var customValidation = new ExcelDataValidationCustom(uid, address, workSheetName);

                    customValidation.Formula.ExcelFormula = oldValidation.As.CustomValidation.Formula.ExcelFormula;

                    return customValidation;
                default:
                    throw new InvalidOperationException("Non supported validationtype: " + oldValidation.ValidationType.Type.ToString());
            }
        }
    }
}
