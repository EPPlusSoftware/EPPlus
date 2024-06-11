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
using OfficeOpenXml.Table.PivotTable.Calculation;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Represents a selection of a row or column field to retreive the calculated value from a pivot table.
    /// </summary>
    public class ExcelPivotTableCalculatedData
    {
        private ExcelPivotTable _pivotTable;
        List<PivotDataFieldItemSelection> _criterias;
        internal ExcelPivotTableCalculatedData(ExcelPivotTable pivotTable)
        {
            _pivotTable = pivotTable;
            _criterias=new List<PivotDataFieldItemSelection>();
        }
        internal ExcelPivotTableCalculatedData(ExcelPivotTable pivotTable, List<PivotDataFieldItemSelection> criterias)
        {
            _pivotTable = pivotTable;
            _criterias = criterias;
        }
        /// <summary>
        /// Specifies which value to use for a field.
        /// </summary>
        /// <param name="fieldName">The name of the field</param>
        /// <param name="value">The value</param>
        /// <returns>A new <see cref="ExcelPivotTableCalculatedData"/> to select other row or column field values or fetch the calulated value in a fluent way.</returns>
        /// <seealso cref="GetValue(string)"/>
        public ExcelPivotTableCalculatedData SelectField(string fieldName, object value)
        {
            CreateField(fieldName, value);
            return new ExcelPivotTableCalculatedData(_pivotTable, _criterias);
        }

        /// <summary>
        /// Specifies which value to use for a field.
        /// </summary>
        /// <param name="fieldName">The name of the field</param>
        /// <param name="value">The value</param>
        /// <param name="subtotalFunction"></param>
        /// <returns></returns>
        public ExcelPivotTableCalculatedData SelectField(string fieldName, object value, eSubTotalFunctions subtotalFunction)
        {
            var fieldSelection = CreateField(fieldName, value); 
            fieldSelection.SubtotalFunction = subtotalFunction;

            _criterias.Add(fieldSelection);
            return new ExcelPivotTableCalculatedData(_pivotTable, _criterias);
        }
        private PivotDataFieldItemSelection CreateField(string fieldName, object value)
        {
            var fieldSelection = new PivotDataFieldItemSelection();
            fieldSelection.FieldName = fieldName;
            fieldSelection.Value = value;
            _criterias.Add(fieldSelection);
            return fieldSelection;
        }

        /// <summary>
        /// Get the value for the current field selection.
        /// <see cref="SelectField(string, object)"/>
        /// <see cref="SelectField(string, object, eSubTotalFunctions)"/>
        /// </summary>
        /// 
        /// <param name="dataFieldName"></param>
        /// <returns></returns>
        public object GetValue(string dataFieldName)
        { 
            return _pivotTable.GetPivotData(dataFieldName, _criterias);
        }
        /// <summary>
        /// Get the value for the current field selection.
        /// <see cref="SelectField(string, object)"/>
        /// <see cref="SelectField(string, object, eSubTotalFunctions)"/>
        /// </summary>
        /// 
        /// <param name="dataFieldIndex">The index for the date field in the <see cref="ExcelPivotTable.DataFields"/> collection</param>
        /// <returns>The value from the pivot table. If data field does not exist of the selected fields does not match any part of the pivot table a #REF! error is retuned.</returns>
        public object GetValue(int dataFieldIndex=0)
        {
            if(dataFieldIndex<0 || dataFieldIndex>=_pivotTable.DataFields.Count)
            {
                return ErrorValues.RefError;
            }
            var name = _pivotTable.DataFields[dataFieldIndex].Name ?? _pivotTable.DataFields[dataFieldIndex].Field.Name;
            return _pivotTable.GetPivotData(name, _criterias);
        }

    }
}