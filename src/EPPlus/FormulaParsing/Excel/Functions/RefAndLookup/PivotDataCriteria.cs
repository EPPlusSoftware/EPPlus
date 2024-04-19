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
    /// <summary>
    /// A criteria for GetPivotData to filter row/column fields
    /// </summary>
    public struct PivotDataCriteria
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fieldName">The row/column field to filter</param>
        /// <param name="value">The value to filter on</param>
        public PivotDataCriteria(string fieldName, object value) 
        {
            FieldName = fieldName;
            Value = value;
        }
        /// <summary>
        /// The row or column field.
        /// </summary>
        public string FieldName { get; set; }
        /// <summary>
        /// The value to filter on.
        /// </summary>
        public object Value { get; set; }
    }
}