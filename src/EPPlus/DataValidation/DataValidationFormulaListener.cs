namespace OfficeOpenXml.DataValidation
{
    //internal class DataValidationFormulaListener : IFormulaListener
    //{
    //    public DataValidationFormulaListener(ExcelDataValidationCollection dataValidations, ExcelWorksheet worksheet)
    //    {
    //        _dataValidations = dataValidations;
    //        _worksheet = worksheet;
    //    }

    //    private ExcelDataValidationCollection _dataValidations;
    //    private ExcelWorksheet _worksheet;

    //    public void Notify(ValidationFormulaChangedArgs e)
    //    {
    //        var validation = _dataValidations.Find(x => x.Uid == e.ValidationUid) as ExcelDataValidation;

    //        if(RefersToOtherWorksheet(e.NewValue))
    //        {
    //            if (validation.InternalValidationType == InternalValidationType.DataValidation)
    //            {
    //                _dataValidations.DataValidationsExt.GetRootNode();
    //                // move from dv to ext
    //                if (validation.ValidationType == ExcelDataValidationType.List)
    //                {
    //                    var listValidation = validation as ExcelDataValidationList;
    //                    var extValidation = new ExcelDataValidationExtList(_worksheet, validation.Uid, validation.Address.AddressSpaceSeparated, validation.ValidationType);
    //                    extValidation.AllowBlank = validation.AllowBlank;
    //                    extValidation.Error = validation.Error;
    //                    extValidation.ErrorStyle = validation.ErrorStyle;
    //                    extValidation.ErrorTitle = validation.ErrorTitle;
    //                    extValidation.ShowErrorMessage = validation.ShowErrorMessage;
    //                    extValidation.HideDropDown = listValidation.ShowDropDown;
    //                    extValidation.Formula.ExcelFormula = e.NewValue;
    //                    listValidation.SetStale();
    //                    _dataValidations.Remove(listValidation);
    //                    _dataValidations.DataValidationsExt.AddValidation(extValidation);
    //                }
    //                else if(validation.ValidationType == ExcelDataValidationType.Custom)
    //                {
    //                    var customValidation = validation as ExcelDataValidationCustom;
    //                    var extValidation = new ExcelDataValidationExtCustom(_worksheet, validation.Uid, validation.Address.AddressSpaceSeparated, validation.ValidationType);
    //                    extValidation.AllowBlank = validation.AllowBlank;
    //                    extValidation.Error = validation.Error;
    //                    extValidation.ErrorStyle = validation.ErrorStyle;
    //                    extValidation.ErrorTitle = validation.ErrorTitle;
    //                    extValidation.ShowErrorMessage = validation.ShowErrorMessage;
    //                    extValidation.Formula.ExcelFormula = e.NewValue;
    //                    customValidation.SetStale();
    //                    _dataValidations.Remove(customValidation);
    //                    _dataValidations.DataValidationsExt.AddValidation(extValidation);
    //                }

    //            }
    //        }
    //        else if(!RefersToOtherWorksheet(e.NewValue))
    //        {
    //            if (validation.InternalValidationType == InternalValidationType.ExtLst)
    //            {
    //                _dataValidations.DataValidationsExt.GetRootNode();
    //                // move from ext to dv
    //                if (validation.ValidationType == ExcelDataValidationType.List)
    //                {
    //                    var listValidation = validation as ExcelDataValidationExtList;
    //                    var dataValidation = _dataValidations.AddListValidation(validation.Address.AddressSpaceSeparated, listValidation.Uid);
    //                    dataValidation.AllowBlank = validation.AllowBlank;
    //                    dataValidation.Error = validation.Error;
    //                    dataValidation.ErrorStyle = validation.ErrorStyle;
    //                    dataValidation.ErrorTitle = validation.ErrorTitle;
    //                    dataValidation.ShowErrorMessage = validation.ShowErrorMessage;
    //                    dataValidation.HideDropDown = listValidation.HideDropDown;
    //                    dataValidation.Formula.ExcelFormula = e.NewValue;
    //                    listValidation.SetStale();
    //                    _dataValidations.DataValidationsExt.Remove(listValidation);
    //                }
    //                else if (validation.ValidationType == ExcelDataValidationType.Custom)
    //                {
    //                    var customValidation = validation as ExcelDataValidationExtCustom;
    //                    _dataValidations.DataValidationsExt.GetRootNode();
    //                    var dataValidation = _dataValidations.AddCustomValidation(validation.Address.AddressSpaceSeparated, customValidation.Uid);
    //                    dataValidation.AllowBlank = validation.AllowBlank;
    //                    dataValidation.Error = validation.Error;
    //                    dataValidation.ErrorStyle = validation.ErrorStyle;
    //                    dataValidation.ErrorTitle = validation.ErrorTitle;
    //                    dataValidation.ShowErrorMessage = validation.ShowErrorMessage;
    //                    dataValidation.Formula.ExcelFormula = e.NewValue;
    //                    customValidation.SetStale();
    //                    _dataValidations.DataValidationsExt.Remove(customValidation);
    //                }
    //            }
    //        }
    //    }

    //    private bool RefersToOtherWorksheet(string address)
    //    {
    //        if(!string.IsNullOrEmpty(address) && ExcelCellBase.IsValidAddress(address))
    //        {
    //            var adr = new ExcelAddress(address);
    //            return !string.IsNullOrEmpty(adr.WorkSheetName) && adr.WorkSheetName != _worksheet.Name;
    //        }
    //        else if(!string.IsNullOrEmpty(address))
    //        {
    //            var tokens = SourceCodeTokenizer.Default.Tokenize(address, _worksheet.Name);
    //            if (!tokens.Any()) return false;
    //            var addressTokens = tokens.Where(x => x.TokenTypeIsSet(TokenType.ExcelAddress));
    //            foreach(var token in addressTokens)
    //            {
    //                var adr = new ExcelAddress(token.Value);
    //                if (!string.IsNullOrEmpty(adr.WorkSheetName) && adr.WorkSheetName != _worksheet.Name)
    //                    return true;
    //            }

    //        }
    //        return false;
    //    }
    //}
}
