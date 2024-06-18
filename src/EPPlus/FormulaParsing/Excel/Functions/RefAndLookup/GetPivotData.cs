/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Calculation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7.2",
        Description = "Returns the value of a pivot table data field.",
        SupportsArrays = false)]
    internal class GetPivotData : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count == 2 && arguments[1].IsExcelRangeOrSingleCell==false)
            {
                //String syntax including a row/column function.
                return GetCriteriasFromString(arguments);
            }
            else
            {
                //Normal syntax
                return GetCriteriasFromArguments(arguments);
            }
        }

        /// <summary>
        /// Gets the Criterias for the row/column field from the normal argument syntax 
        /// </summary>
        /// <param name="arguments">The arguments to the GetPivotData</param>
        /// <returns>The compiled result</returns>
        private CompileResult GetCriteriasFromString(IList<FunctionArgument> arguments)
        {
            var address = arguments[0].ValueAsRangeInfo;
            if (address == null)
            {
                return CompileResult.GetErrorResult(eErrorType.Name);
            }
            if (GetPivotTable(address, out ExcelPivotTable pivotTable) == false)
            {
                return CompileResult.GetErrorResult(eErrorType.Name);
            }

            string dataFieldName = "";
            List<string> fieldValues = new List<string>();
            var criteria = new List<PivotDataFieldItemSelection>();
            var criteriaString = arguments[1].Value.ToString();
            var sb = new StringBuilder();
            int bracketCount = 0;
            bool isInString = false;
            bool hasValue = false;
            bool? functionIsRowField = null;
            for (int i = 0;i< criteriaString.Length;i++)
            {
                var c = criteriaString[i];
                if (c == ' ' && isInString == false)
                {
                    if (sb.Length > 0)
                    {
                        fieldValues.Add(sb.ToString());
                        sb.Length=0;
                    }
                }
                else if (c == '\'')
                {
                    isInString = !isInString;
                    if (i > 0 && criteriaString[i - 1] == '\'')
                    {
                        sb.Append(c);
                    }
                }
                else if (c == '[' && isInString == false)
                {
                    var fieldName = sb.ToString();
                    fieldValues.Add(fieldName);
                    sb.Length = 0;
                    var f = pivotTable.Fields[fieldName];
                    if(f==null || (f.IsRowField==false && f.IsColumnField==false))
                    {
                        return CompileResult.GetErrorResult(eErrorType.Ref);
                    }
                    else if (f.IsRowField)
                    {
                        functionIsRowField = true;
                    }
                    else 
                    {
                        functionIsRowField = false;
                    }
                    AddCriterias(fieldValues, (functionIsRowField.Value ? pivotTable.RowFields: pivotTable.ColumnFields), ref criteria);
                    fieldValues.Clear();
                    bracketCount++;
                }
                else if ((c == ','  || c == ';') && isInString == false && bracketCount>0)
                {
                    if(hasValue)
                    {
                        return CompileResult.GetErrorResult(eErrorType.Ref);
                    }
                    criteria[criteria.Count-1].Value = sb.ToString();
                    hasValue = true;
                    sb = new StringBuilder();
                }
                else if (c ==']' && isInString == false)
                {
                    if(GetSubTotalFunctionFromString(sb.ToString(), out eSubTotalFunctions function))
                    {
                        criteria[criteria.Count - 1].SubtotalFunction = function;
                    }
                    sb = new StringBuilder();
                    bracketCount--;
                }
                else
                {
                    sb.Append(c);
                }
            }

            if (sb.Length > 0)
            {
                var fieldName = sb.ToString();
                var field = pivotTable.Fields[fieldName];
                if (field != null && field.IsDataField)
                {
                    dataFieldName = fieldName;
                }
                else
                {
                    fieldValues.Add(fieldName);
                }
            }

            if(fieldValues.Count > 0)
            {
                if (functionIsRowField.HasValue)
                {
                    AddCriterias(fieldValues, (functionIsRowField.Value ? pivotTable.ColumnFields : pivotTable.RowFields), ref criteria);
                    criteria[criteria.Count - 1].Value = fieldValues[fieldValues.Count - 1];
                }
                else
                {
                    return CompileResult.GetErrorResult(eErrorType.Ref);
                }
            }

            if(dataFieldName==string.Empty && pivotTable.DataFields.Count==1)
            {
                dataFieldName = pivotTable.DataFields[0].Name;
                if (dataFieldName == string.Empty) dataFieldName = pivotTable.DataFields[0].Field.Name;
            }

            var result = pivotTable.GetPivotData(dataFieldName, criteria);
            return new CompileResult(result, DataType.Decimal);
        }

        private void AddCriterias(List<string> fieldValues, ExcelPivotTableRowColumnFieldCollection fields, ref List<PivotDataFieldItemSelection> criteria)
        {
            for (int i=0;i<fieldValues.Count-1; i++)
            {
                criteria.Add(new PivotDataFieldItemSelection(fields[i].Name, fieldValues[i]));
            }
            criteria.Add(new PivotDataFieldItemSelection(fields[fieldValues.Count - 1].Name, null));
        }

        private static string AddFieldValue(ExcelPivotTable pivotTable, List<PivotDataFieldItemSelection> criteria, ref StringBuilder sb, int rowColIndex)
        {
            string fieldName = "";
            if (rowColIndex < pivotTable.RowFields.Count)
            {
                fieldName = pivotTable.RowFields[rowColIndex].Name;
            }
            else
            {
                var ix = rowColIndex - pivotTable.RowFields.Count;
                if (ix < pivotTable.ColumnFields.Count)
                {
                    fieldName = pivotTable.ColumnFields[ix].Name;
                }
                else
                {
                    var dataFieldName = sb.ToString();
                    sb = new StringBuilder();
                    return dataFieldName;
                }
            }
            var sel = new PivotDataFieldItemSelection();
            sel.FieldName = fieldName;
            sel.Value = sb.ToString();
            criteria.Add(sel);
            sb = new StringBuilder();
            return "";
        }

        private bool GetSubTotalFunctionFromString(string value, out eSubTotalFunctions function)
        {
            switch (value.ToLower())
            {
                case "sum":
                    function = eSubTotalFunctions.Sum;
                    break;
                case "count":
                    function = eSubTotalFunctions.CountA;
                    break;
                case "count nums":
                    function = eSubTotalFunctions.Count;
                    break;
                case "average":
                    function = eSubTotalFunctions.Avg;
                    break;
                case "min":
                    function = eSubTotalFunctions.Min;
                    break;
                case "max":
                    function = eSubTotalFunctions.Max;
                    break;
                case "stddev":
                    function = eSubTotalFunctions.StdDev;
                    break;
                case "stddevp":
                    function = eSubTotalFunctions.StdDevP;
                    break;
                case "var":
                    function = eSubTotalFunctions.Var;
                    break;
                case "varp":
                    function = eSubTotalFunctions.VarP;
                    break;
                case "product":
                    function = eSubTotalFunctions.Product;
                    break;
                default:
                    function = eSubTotalFunctions.None;
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Gets the Criterias a string. This syntax is used when a row/column field has its own subtotals. 
        /// In this case the first parameter is the address to the pivot table and the second parameter is a string containing all information regarding criteria and which function is used.
        /// Syntax 'Field Name'['Field Value',Function]. If the value is not the first row/column field values are space separated before and after. Example =GETPIVOTDATA($B$2;"Australia Sindey 'Years (InvoiceDate)'['2022',Count] '9232'") .
        /// </summary>
        /// <param name="arguments">The arguments to the GetPivotData</param>
        /// <returns>The compiled result</returns>
        private CompileResult GetCriteriasFromArguments(IList<FunctionArgument> arguments)
        {
            var address = arguments[1].ValueAsRangeInfo;
            if (address == null)
            {
                return CompileResult.GetErrorResult(eErrorType.Name);
            }
            if (GetPivotTable(address, out ExcelPivotTable pivotTable) == false)
            {
                return CompileResult.GetErrorResult(eErrorType.Name);
            }
            var dataFieldName = (arguments[0].Value ?? "").ToString();

            var dataField = pivotTable.DataFields[dataFieldName];
            if (dataField == null)
            {
                return CompileResult.GetErrorResult(eErrorType.Ref);
            }

            var criteria = new List<PivotDataFieldItemSelection>();
            for (int i = 2; i < arguments.Count; i += 2)
            {
                var field = pivotTable.Fields[ArgToString(arguments, i)];
                if (field == null)
                {
                    return CompileResult.GetErrorResult(eErrorType.Ref);
                }
                var value = arguments[i + 1].Value;

                criteria.Add(new PivotDataFieldItemSelection(field.Name, value));
            }

            //Calulate value;      
            var result = pivotTable.GetPivotData(dataFieldName, criteria);
            return new CompileResult(result, DataType.Decimal);
        }

        private bool GetPivotTable(IRangeInfo ri, out ExcelPivotTable pivotTable)
        {
            var ws = ri.Worksheet;
            var adr = ri.Address;
            foreach (var pt in ws.PivotTables)
            {
                if (pt.Address.Collide(adr.FromRow, adr.FromCol, adr.ToRow, adr.ToCol)!=ExcelAddressBase.eAddressCollition.No)
                {
                    pivotTable = pt;
                    return true;
                }
            }
            pivotTable = null;
            return false;
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
