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
        EPPlusVersion = "7",
        Description = "Returns the value of a pivot table data field.",
        SupportsArrays = false)]
    internal class GetPivotData : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count == 2 && arguments[1].IsExcelRangeOrSingleCell==false)
            {
                return GetCriteriasTokenizedString(arguments);
            }
            else
            {
                return GetCriteriasArguments(arguments);
            }
        }

        private CompileResult GetCriteriasTokenizedString(IList<FunctionArgument> arguments)
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
            var criteria = new List<PivotDataFieldItemSelection>();
            var criteriaString = arguments[1].Value.ToString();
            var sb = new StringBuilder();
            int bracketCount = 0;
            bool isInString = false;
            int rowColIndex = 0;
            for (int i = 0;i< criteriaString.Length;i++)
            {
                var c = criteriaString[i];
                if (c == ' ' && isInString == false)
                {
                    if (sb.Length > 0)
                    {
                        dataFieldName = AddFieldValue(pivotTable, criteria, ref sb, rowColIndex);
                        rowColIndex++;
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
                    var sel = new PivotDataFieldItemSelection();
                    sel.FieldName = sb.ToString();
                    criteria.Add(sel);
                    sb = new StringBuilder();
                    bracketCount++;
                    if(rowColIndex < pivotTable.RowFields.Count)
                    {
                        rowColIndex = pivotTable.RowFields.Count;
                    }
                    else
                    {
                        rowColIndex = pivotTable.RowFields.Count+ pivotTable.ColumnFields.Count;
                    }
                }
                else if ((c == ','  || c == ';') && isInString == false && bracketCount>0)
                {
                    criteria[criteria.Count-1].Value = sb.ToString();
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
                dataFieldName = AddFieldValue(pivotTable, criteria, ref sb, rowColIndex);
            }
            
            if(dataFieldName==string.Empty && pivotTable.DataFields.Count==1)
            {
                dataFieldName = pivotTable.DataFields[0].Name;
                if (dataFieldName == string.Empty) dataFieldName = pivotTable.DataFields[0].Field.Name;
            }
            var result = pivotTable.GetPivotData(dataFieldName, criteria);
            return new CompileResult(result, DataType.Decimal);
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

        private CompileResult GetCriteriasArguments(IList<FunctionArgument> arguments)
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
