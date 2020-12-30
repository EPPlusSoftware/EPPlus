using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableStyle
    {
        ExcelStyles _styles;
        internal ExcelPivotTableStyle(ExcelPivotTable pt)
        {
            _styles = pt.WorkSheet.Workbook.Styles;
            Styles = new ExcelPivotTableAreaStyleCollection(pt);
        }
        //ExcelPivotArea _all = null;
        //public ExcelPivotArea All 
        //{ 
        //    get
        //    {

        //    }
        //}
        public ExcelPivotTableAreaStyleCollection Styles
        {
            get;
        } 
    }
}
