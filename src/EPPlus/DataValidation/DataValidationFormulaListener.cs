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
                    var extValidation = new ExcelDataValidationExtList(_worksheet, ExcelDataValidation.NewId(), validation.Address.Address, validation.ValidationType);
                    extValidation.Formula.ExcelFormula = listValidation.Formula.ExcelFormula;
                    _dataValidations.DataValidationsExt.AddValidation(extValidation);
                    _dataValidations.Remove(listValidation);
                }
                
            }
            else if(validation.InternalValidationType == InternalValidationType.ExtLst
                &&
                !RefersToOtherWorksheet(e.NewValue))
            {
                // move from ext to dv
            }
        }

        private bool RefersToOtherWorksheet(string address)
        {
            var adr = new ExcelAddress(address);
            return string.IsNullOrEmpty(adr.WorkSheet) || adr.WorkSheet != _worksheet.Name;
        }
    }
}
