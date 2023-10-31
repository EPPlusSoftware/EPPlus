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
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public struct PivotDataCriteria
    {
        public PivotDataCriteria(ExcelPivotTableField field, object value) 
        {
            Field = field;
            Value = value;
        }
        public ExcelPivotTableField Field { get; set; }
        public object Value { get; set; }
    }
}