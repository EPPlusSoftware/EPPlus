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

        //T AddValidation<T>(string address, ExcelDataValidationType ValidationType, Type type)
        //where T : IExcelDataValidation
        //{
        //}

        //internal static T Create<T>(string address, ExcelDataValidationType ValidationType, T method)
        //    where T : NewDataValidation
        //{
        //    object item = Activator.CreateInstance(typeof(T), address, ExcelDataValidation.NewId(), ValidationType);
        //    return (T)item;
        //}
        internal static ExcelDataValidation Create(XmlReader xr)
        {
            string validationTypeName = xr.GetAttribute("type");
            ExcelDataValidation alt;
            switch (validationTypeName)
            {
                case null:
                    alt = new ExcelDataValidationAny(xr);
                    break;
                case "textLength":
                case "whole":
                    return new ExcelDataValidationInt(xr);
                case "decimal":
                    return new ExcelDataValidationDecimal(xr);
                case "list":
                    return new ExcelDataValidationList(xr);
                //case eDataValidationType.DateTime:
                //    return new ExcelDataValidationDateTime(worksheet, uid, address, type, itemElementNode);
                //case eDataValidationType.Time:
                //    return new ExcelDataValidationTime(worksheet, uid, address, type, itemElementNode);
                //case eDataValidationType.Custom:
                //    return CreateCustomValidation(type, worksheet, address, itemElementNode, internalType, uid);
                default:
                    throw new InvalidOperationException($"Non supported validationtype: {validationTypeName}");


            }

            alt.LoadXML(xr);

            return null;
        }

        internal static ExcelDataValidation Create(ExcelDataValidationType type, string address, string uid)
        {
            switch (type.Type)
            {
                case eDataValidationType.Any:
                    return new ExcelDataValidationAny(uid, address);
                default:
                    throw new InvalidOperationException($"Non supported validationtype: {type}");
            }
            return null;
        }


        //internal static ExcelDataValidation Create(string uid, string address, string validationType)
        //{

        //    //switch (validationType.Type)
        //    //{
        //    //    case eDataValidationType.Any:
        //    //        return new ExcelDataValidationAny(uid, address, validationType);
        //    //    //case eDataValidationType.TextLength:
        //    //    //case eDataValidationType.Whole:
        //    //    //    return new ExcelDataValidationInt(worksheet, uid, address, type, itemElementNode);
        //    //    //case eDataValidationType.Decimal:
        //    //    //    return new ExcelDataValidationDecimal(worksheet, uid, address, type, itemElementNode);
        //    //    //case eDataValidationType.List:
        //    //    //    return CreateListValidation(type, worksheet, address, itemElementNode, internalType, uid);
        //    //    //case eDataValidationType.DateTime:
        //    //    //    return new ExcelDataValidationDateTime(worksheet, uid, address, type, itemElementNode);
        //    //    //case eDataValidationType.Time:
        //    //    //    return new ExcelDataValidationTime(worksheet, uid, address, type, itemElementNode);
        //    //    //case eDataValidationType.Custom:
        //    //    //    return CreateCustomValidation(type, worksheet, address, itemElementNode, internalType, uid);
        //    //    default:
        //    //        throw new InvalidOperationException("Non supported validationtype: " + validationType.Type.ToString());
        //    //}
        //    //object item = Activator.CreateInstance(typeof(T), address, ExcelDataValidation.NewId(), InternalValidationType internalType);
        //    //return (T)item;
        //    return null;
        //}


        //internal static ExcelDataValidation Create(ExcelDataValidationType type, ExcelWorksheet worksheet, string address, XmlNode itemElementNode, InternalValidationType internalType, string uid)
        //{
        //    Require.Argument(type).IsNotNull("validationType");
        //    switch (type.Type)
        //    {
        //        case eDataValidationType.Any:
        //            return new ExcelDataValidationAny(worksheet, uid, address, type, itemElementNode);
        //        case eDataValidationType.TextLength:
        //        case eDataValidationType.Whole:
        //            return new ExcelDataValidationInt(worksheet, uid, address, type, itemElementNode);
        //        case eDataValidationType.Decimal:
        //            return new ExcelDataValidationDecimal(worksheet, uid, address, type, itemElementNode);
        //        case eDataValidationType.List:
        //            return CreateListValidation(type, worksheet, address, itemElementNode, internalType, uid);
        //        case eDataValidationType.DateTime:
        //            return new ExcelDataValidationDateTime(worksheet, uid, address, type, itemElementNode);
        //        case eDataValidationType.Time:
        //            return new ExcelDataValidationTime(worksheet, uid, address, type, itemElementNode);
        //        case eDataValidationType.Custom:
        //            return CreateCustomValidation(type, worksheet, address, itemElementNode, internalType, uid);
        //        default:
        //            throw new InvalidOperationException("Non supported validationtype: " + type.Type.ToString());
        //    }
        //}

        internal static ExcelDataValidationWithFormula<IExcelDataValidationFormulaList> CreateListValidation(string address, string uid, string workSheetName)
        {
            //if (internalType == InternalValidationType.DataValidation)
            //{
            return new ExcelDataValidationList(uid, address, workSheetName);
            //}

            //// extLst
            //return new ExcelDataValidationExtList(worksheet, uid, address, type, itemElementNode);
        }

        internal static ExcelDataValidationWithFormula<IExcelDataValidationFormula> CreateCustomValidation(string address, string uid, string workSheetName)
        {
            //if (internalType == InternalValidationType.DataValidation)
            //{
            return new ExcelDataValidationCustom(uid, address, workSheetName);
            // }
            //// extLst
            //return new ExcelDataValidationExtCustom(worksheet, uid, address, type, itemElementNode);
        }
    }
}
