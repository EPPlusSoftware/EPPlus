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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.DataValidation
{
    internal class RangeDataValidation : IRangeDataValidation
    {
        public RangeDataValidation(ExcelWorksheet worksheet, string address)
        {
            Require.Argument(worksheet).IsNotNull("worksheet");
            Require.Argument(address).IsNotNullOrEmpty("address");
            _worksheet = worksheet;
            _address = address;
        }

        ExcelWorksheet _worksheet;
        string _address;

        /// <summary>
        ///  Used to remove all dataValidations in cell or cellrange
        /// </summary>
        /// <param name="deleteIfEmpty">Deletes the dataValidation if it has no addresses after clear</param>
        /// <exception cref="InvalidOperationException"></exception>
        public void ClearDataValidation(bool deleteIfEmpty = false)
        {
            var address = new ExcelAddress(_address);
            var validations = _worksheet.DataValidations._validationsRD.GetValuesFromRange(address._fromRow, address._fromCol, address._toRow, address._toCol);

            foreach( var validation in validations)
            {
                var excelAddress = new ExcelAddressBase(validation.Address.Address.Replace(" ", ","));
                var addresses = excelAddress.GetAllAddresses();

                string newAddress = "";

                foreach (var validationAddress in addresses)
                {
                    var nullOrAddress = validationAddress.IntersectReversed(address);
                    
                    if (nullOrAddress != null)
                    {
                        newAddress+= nullOrAddress.Address + " ";
                    }
                }

                if (newAddress == "")
                {
                    if (deleteIfEmpty)
                    {
                        _worksheet.DataValidations.Remove(validation);
                    }
                    else
                    {
                        throw new InvalidOperationException($"Cannot remove last address in validation of type {validation.ValidationType.Type} " +
                            $"with uid {validation.Uid} without deleting it." +
                            $" Add other addresses or use ClearDataValidation(true)");
                    }
                }
                else
                {
                    validation.Address.Address = newAddress;
                }
            }
        }

        public IExcelDataValidationAny AddAnyDataValidation()
        {
            return _worksheet.DataValidations.AddAnyValidation(_address);
        }

        public Contracts.IExcelDataValidationInt AddIntegerDataValidation()
        {
            return _worksheet.DataValidations.AddIntegerValidation(_address);
        }

        public IExcelDataValidationDecimal AddDecimalDataValidation()
        {
            return _worksheet.DataValidations.AddDecimalValidation(_address);
        }

        public IExcelDataValidationDateTime AddDateTimeDataValidation()
        {
            return _worksheet.DataValidations.AddDateTimeValidation(_address);
        }

        public IExcelDataValidationList AddListDataValidation()
        {
            return _worksheet.DataValidations.AddListValidation(_address);
        }

        public Contracts.IExcelDataValidationInt AddTextLengthDataValidation()
        {
            return _worksheet.DataValidations.AddTextLengthValidation(_address);
        }

        public IExcelDataValidationTime AddTimeDataValidation()
        {
            return _worksheet.DataValidations.AddTimeValidation(_address);
        }

        public IExcelDataValidationCustom AddCustomDataValidation()
        {
            return _worksheet.DataValidations.AddCustomValidation(_address);
        }
    }
}
