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
using System.Net;

namespace OfficeOpenXml.DataValidation
{
    internal class RangeDataValidation : IRangeDataValidation
    {
        public RangeDataValidation(ExcelWorksheet worksheet, ExcelAddress address)
        {
            Require.Argument(worksheet).IsNotNull("worksheet");
            Require.Argument(address.Address).IsNotNullOrEmpty("address");
            _worksheet = worksheet;
            _address = address;
        }

        ExcelWorksheet _worksheet;
        ExcelAddress _address;

        /// <summary>
        ///  Used to remove all dataValidations in cell or cellrange
        /// </summary>
        /// <param name="deleteIfEmpty">Deletes the dataValidation if it has no addresses after clear</param>
        /// <exception cref="InvalidOperationException"></exception>
        public void ClearDataValidation(bool deleteIfEmpty = false)
        {
            var validations = _worksheet.DataValidations._validationsRD.GetValuesFromRange(_address._fromRow, _address._fromCol, _address._toRow, _address._toCol);

            foreach( var validation in validations)
            {
                var excelAddress = new ExcelAddressBase(validation.Address.Address.Replace(" ", ","));
                var addresses = excelAddress.GetAllAddresses();

                string newAddress = "";

                foreach (var validationAddress in addresses)
                {
                    var nullOrAddress = validationAddress.IntersectReversed(_address);
                    
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
            return _worksheet.DataValidations.AddAnyValidation(_address.Address);
        }

        public Contracts.IExcelDataValidationInt AddIntegerDataValidation()
        {
            return _worksheet.DataValidations.AddIntegerValidation(_address.Address);
        }

        public IExcelDataValidationDecimal AddDecimalDataValidation()
        {
            return _worksheet.DataValidations.AddDecimalValidation(_address.Address);
        }

        public IExcelDataValidationDateTime AddDateTimeDataValidation()
        {
            return _worksheet.DataValidations.AddDateTimeValidation(_address.Address);
        }

        public IExcelDataValidationList AddListDataValidation()
        {
            return _worksheet.DataValidations.AddListValidation(_address.Address);
        }

        public Contracts.IExcelDataValidationInt AddTextLengthDataValidation()
        {
            return _worksheet.DataValidations.AddTextLengthValidation(_address.Address);
        }

        public IExcelDataValidationTime AddTimeDataValidation()
        {
            return _worksheet.DataValidations.AddTimeValidation(_address.Address);
        }

        public IExcelDataValidationCustom AddCustomDataValidation()
        {
            return _worksheet.DataValidations.AddCustomValidation(_address.Address);
        }

        public List<ExcelDataValidation> GetDataValidations()
        {
            var hs = new HashSet<ExcelDataValidation>();
            var l = _worksheet.DataValidations.GetIntersectingRanges(_address);
            foreach (var i in l)
            {
                var v = (ExcelDataValidation)i.Value;
                if (!hs.Contains(v))
                {
                    hs.Add(v);
                }
            }
            return hs.ToList();
        }
    }
}
