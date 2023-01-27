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
using OfficeOpenXml.Utils;

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

        public IExcelDataValidationAny AddAnyDataValidation()
        {
            return _worksheet.DataValidations.AddAnyValidation(_address);
        }

        public IExcelDataValidationInt AddIntegerDataValidation()
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

        public IExcelDataValidationInt AddTextLengthDataValidation()
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
