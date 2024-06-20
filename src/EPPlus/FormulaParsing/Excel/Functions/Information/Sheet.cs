using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Information,
        EPPlusVersion = "5.5",
        IntroducedInExcelVersion = "2013",
        Description = "Returns the sheet number relating to a supplied reference")]
    internal class Sheet : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 0;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var result = -1;
            if(arguments.Count == 0)
            {
                result = context.CurrentCell.WorksheetIx + 1;
            }
            else
            {
                var arg = arguments[0];
                if(arg.Address!=null)
                {
                    if (arg.Address.WorksheetIx>=0)
                    {
                        result = arg.Address.WorksheetIx + 1;
                        //var excelAddress = new ExcelAddress(address);
                        //result = context.ExcelDataProvider.GetWorksheetIndex(excelAddress.WorkSheetName);
                    }
                    else
                    {
                        var address = ArgToAddress(arguments, 0);
                        var value = string.IsNullOrEmpty(address) ? ArgToString(arguments, 0) : address;
                        var worksheetNames = context.ExcelDataProvider.GetWorksheets();
                        
                        // for each worksheet in the workbook - check if the value a worksheet name.
                        foreach(var wsName in worksheetNames)
                        {
                            if(string.Compare(wsName, value, true) == 0)
                            {
                                result = context.ExcelDataProvider.GetWorksheetIndex(wsName);
                                break;
                            }
                        }
                        if (result == -1)
                        {
                            // not a worksheet name, now check if it is a named range in the current worksheet
                            var wsNamedRanges = context.CurrentWorksheet.Names;
                            var matchingWsName = wsNamedRanges.FirstOrDefault(x => x.Name == value);
                            if (matchingWsName != null)
                            {
                                result = context.ExcelDataProvider.GetWorksheetIndex(matchingWsName.WorkSheetName);
                            }

                            if (result == -1)
                            {
                                // not a worksheet named range, now check workbook level
                                var namedRanges = context.ExcelDataProvider.GetWorkbookNameValues();
                                var matchingWorkbookRange = namedRanges.FirstOrDefault(x => x.Name == value);
                                if (matchingWorkbookRange != null)
                                {
                                    result = context.ExcelDataProvider.GetWorksheetIndex(matchingWorkbookRange.WorkSheetName);
                                }
                                else
                                {
                                    result = context.ExcelDataProvider.GetWorksheetIndex(value);
                                }
                            }

                            if (result == -1)
                            {
                                var table = context.ExcelDataProvider.GetExcelTable(value);
                                if (table != null)
                                {
                                    result = context.ExcelDataProvider.GetWorksheetIndex(table.WorkSheet.Name);
                                }
                            }
                        }
                    }
                }
                else
                {
                    var value = ArgToString(arguments, 0);
                    result = context.ExcelDataProvider.GetWorksheetIndex(value);
                }
            }
            if(result == -1)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            return CreateResult(result, DataType.Integer);
        }
        /// <summary>
        /// Reference Parameters do not need to be follows in the dependency chain.
        /// </summary>
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreAddress;
        }));
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
