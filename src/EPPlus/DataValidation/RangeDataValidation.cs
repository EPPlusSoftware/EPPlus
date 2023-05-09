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

        private List<ExcelAddress> GetAllAddresses(ExcelAddress addressBase)
        {
            if(addressBase.Addresses == null)
            {
                return new List<ExcelAddress> { addressBase };
            }
            else 
            {
                return addressBase.Addresses;
            }
        }

        public void ClearDataValidation(bool deleteIfEmpty = false)
        {
            var adress = new ExcelAddress(_address);

            GetAllAddresses(adress).ForEach(a =>
            {
                var validation = _worksheet.DataValidations.
                Find(x => x.Address.Collide(a) != ExcelAddressBase.eAddressCollition.No);

                if (validation != null)
                {
                    var addresses = GetAllAddresses(validation.Address);

                    if(addresses.Count == 1)
                    {
                        //ExcelRangeBase test = _worksheet.Cells[addresses[0].Address];

                        var address2 = addresses[0].IntersectReversed(a);

                        validation.Address.Address = address2.Address;

                        //validation.Address.Address = ;
                        //ExcelRange range = new ExcelRange(_worksheet, addresses[0].Address);
                        //foreach(ExcelAddress address in range) 
                        //{
                        //    if(address == a)
                        //    {
                        //        range.Except(a.Address);
                        //    }
                        //}
                    }
                    else
                    {
                        validation.Address.Addresses.Remove(a);
                    }
                    //validation.Address.Address;
                    //var addresses = GetAllAddresses(validation.Address);
                    //validation.Address.Addresses.Find(va => va == a);
                }
            });

            //for (int i = 0; i< adress.Addresses.Count; i++)
            //{

            //}

            //var validation = _worksheet.DataValidations.
            //    Find(x => x.Address.Collide(adress.Addresses) != ExcelAddressBase.eAddressCollition.No);

            //var newString = validation.Address.Address;

            ////trigger the setter
            //validation.Address.Address = newString.Replace(_address, "");

            //if (deleteIfEmpty)
            //{
            //    if (validation.Address.Address == "")
            //    {
            //        _worksheet.DataValidations.Remove(validation);
            //    }
            //}
            //else
            //{
            //    throw new InvalidOperationException($"All addresses within validation {validation} were removed. " +
            //        $"Please leave at least one address or call ClearDataValidation(true) instead.");
            //}

            //_worksheet.DataValidations.DeleteRangeDictionary(new ExcelAddress(_address), false);
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
