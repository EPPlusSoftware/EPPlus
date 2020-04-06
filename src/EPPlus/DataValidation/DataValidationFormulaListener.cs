using OfficeOpenXml.DataValidation.Events;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    internal class DataValidationFormulaListener : IFormulaListener
    {
        public DataValidationFormulaListener(ExcelDataValidationCollection dataValidations, ExcelWorksheet worksheet)
        {
            _dataValidations = dataValidations;
            _worksheet = worksheet;
        }

        private ExcelDataValidationCollection _dataValidations;
        private ExcelWorksheet _worksheet;

        public void Notify(ValidationFormulaChangedArgs e)
        {
            var validation = _dataValidations.Find(x => x.Uid == e.ValidationUid) as ExcelDataValidation;
            if(validation.InternalValidationType == InternalValidationType.DataValidation
                &&
                RefersToOtherWorksheet(e.NewValue))
            {
                // move from dv to ext
                if(validation.ValidationType == ExcelDataValidationType.List)
                {
                    var listValidation = validation as ExcelDataValidationList;
                    _dataValidations.DataValidationsExt.GetRootNode();
                    var extValidation = new ExcelDataValidationExtList(_worksheet, validation.Uid, validation.Address.Address, validation.ValidationType);
                    extValidation.AllowBlank = validation.AllowBlank;
                    extValidation.Error = validation.Error;
                    extValidation.ErrorStyle = validation.ErrorStyle;
                    extValidation.ErrorTitle = validation.ErrorTitle;
                    extValidation.ShowErrorMessage = validation.ShowErrorMessage;
                    extValidation.Formula.ExcelFormula = e.NewValue;
                    _dataValidations.Remove(listValidation);
                    _dataValidations.DataValidationsExt.AddValidation(extValidation);
                }
                
            }
            else if(validation.InternalValidationType == InternalValidationType.ExtLst
                &&
                !RefersToOtherWorksheet(e.NewValue))
            {
                // move from ext to dv
                if (validation.ValidationType == ExcelDataValidationType.List)
                {
                    var listValidation = validation as ExcelDataValidationExtList;
                    _dataValidations.DataValidationsExt.GetRootNode();
                    var dataValidation = _dataValidations.AddListValidation(validation.Address.Address, listValidation.Uid);
                    dataValidation.AllowBlank = validation.AllowBlank;
                    dataValidation.Error = validation.Error;
                    dataValidation.ErrorStyle = validation.ErrorStyle;
                    dataValidation.ErrorTitle = validation.ErrorTitle;
                    dataValidation.ShowErrorMessage = validation.ShowErrorMessage;
                    dataValidation.Formula.ExcelFormula = e.NewValue;
                    _dataValidations.DataValidationsExt.Remove(listValidation);
                }
            }
        }

        private bool RefersToOtherWorksheet(string address)
        {
            if(ExcelAddressBase.IsValidAddress(address))
            {
                var adr = new ExcelAddress(address);
                return !string.IsNullOrEmpty(adr.WorkSheetName) && adr.WorkSheetName != _worksheet.Name;
            }
            else
            {
                return false;
            }
        }
    }
}
