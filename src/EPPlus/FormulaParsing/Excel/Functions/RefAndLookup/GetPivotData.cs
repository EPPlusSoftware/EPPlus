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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Table.PivotTable;
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

            var dataField = pivotTable.DataFields.FirstOrDefault(x=>x.Field.Name==dataFieldName);
            if(dataField == null)
            {
                return CompileResult.GetErrorResult(eErrorType.Ref);
            }

            var criteria = new List<PivotDataCriteria>();
            for(int i=2;i<arguments.Count;i+=2) 
            { 
                var field = pivotTable.Fields[ArgToString(arguments,i)];
                if (field == null)
                {
                    return CompileResult.GetErrorResult(eErrorType.Ref);
                }
                var value = arguments[i+1].Value;

                criteria.Add(new PivotDataCriteria(field, value));
            }
            //Calulate value;      
            var result = pivotTable.GetPivotData(criteria, dataField);
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
    }
}
