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
        /// Creates an instance of <see cref="ExcelDataValidation"/> out of the reader.
        /// </summary>
        /// <param name="xr"></param>
        /// <returns>"</returns>
        /// <exception cref="InvalidOperationException"></exception>
        internal static ExcelDataValidation Create(XmlReader xr, ExcelWorksheet ws)
        {
            string validationTypeName = xr.GetAttribute("type") == null ? "" : xr.GetAttribute("type");

            switch (validationTypeName)
            {
                case "":
                    return new ExcelDataValidationAny(xr, ws);
                case "textLength":
                    return new ExcelDataValidationInt(xr, ws, true);
                case "whole":
                    return new ExcelDataValidationInt(xr, ws);
                case "decimal":
                    return new ExcelDataValidationDecimal(xr, ws);
                case "list":
                    return new ExcelDataValidationList(xr, ws);
                case "time":
                    return new ExcelDataValidationTime(xr, ws);
                case "date":
                    return new ExcelDataValidationDateTime(xr, ws);
                case "custom":
                    return new ExcelDataValidationCustom(xr, ws);
                default:
                    throw new InvalidOperationException($"Non supported validationtype: {validationTypeName}");
            }
        }

        static internal ExcelDataValidation CloneWithNewAdress(string address, ExcelDataValidation oldValidation, ExcelWorksheet added)
        {
            var validation = oldValidation.GetClone(added);
            validation.Address = new ExcelDatavalidationAddress(address, validation);
            return validation;
        }
    }
}
