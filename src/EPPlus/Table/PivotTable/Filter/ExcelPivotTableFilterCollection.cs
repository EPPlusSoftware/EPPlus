/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Xml;

namespace EPPlusTest.Table.PivotTable.Filter
{
    /// <summary>
    /// A collection of pivot filters for a pivot table
    /// </summary>
    public class ExcelPivotTableFilterCollection : ExcelPivotTableFilterBaseCollection
    {
        internal ExcelPivotTableFilterCollection(ExcelPivotTable table) : base(table)
        {
        }
    }
}
