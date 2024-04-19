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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableCalculatedData
    {
        private ExcelPivotTable _pivotTable;
        List<PivotDataCriteria> _criterias;
        internal ExcelPivotTableCalculatedData(ExcelPivotTable pivotTable)
        {
            _pivotTable = pivotTable;
            _criterias=new List<PivotDataCriteria>();
        }
        internal ExcelPivotTableCalculatedData(ExcelPivotTable pivotTable, List<PivotDataCriteria> criterias)
        {
            _pivotTable = pivotTable;
            _criterias = criterias;
        }
        public ExcelPivotTableCalculatedData Criterias(Action<PivotDataCriteria> x)
        {
            var criteria = new PivotDataCriteria();
            x.Invoke(criteria);
            _criterias.Add(criteria);
            return new ExcelPivotTableCalculatedData(_pivotTable, _criterias);
        }
        public object GetValue(string dataFieldName)
        { 
            return _pivotTable.GetPivotData(dataFieldName, _criterias);
        }
    } 
}