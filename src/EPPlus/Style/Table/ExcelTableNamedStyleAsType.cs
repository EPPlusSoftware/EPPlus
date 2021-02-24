/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Style.Table
{
    /// <summary>
    /// Provides a simple way to type cast a table named style objects to its top level class.
    /// </summary>
    public class ExcelTableNamedStyleAsType
    {
        ExcelTableNamedStyleBase _tableNamedStyle;
        internal ExcelTableNamedStyleAsType(ExcelTableNamedStyleBase tableNamedStyle)
        {
            _tableNamedStyle = tableNamedStyle;
        }

        /// <summary>
        /// Converts the table named style object to it's top level or another nested class.        
        /// </summary>
        /// <typeparam name="T">The type of table named style object. T must be inherited from ExcelTableNamedStyleBase</typeparam>
        /// <returns>The table named style as type T</returns>
        public T Type<T>() where T : ExcelTableNamedStyleBase
        {
            if(_tableNamedStyle is T t)
            {
                return t;
            }
            return default;
        }
        /// <summary>
        /// Returns the table named style object as a named style for tables only
        /// </summary>
        /// <returns>The table named style object</returns>
        public ExcelTableNamedStyle TableStyle
        {
            get
            {
                return _tableNamedStyle as ExcelTableNamedStyle;
            }
        }
        /// <summary>
        /// Returns the table named style object as a named style for pivot tables only
        /// </summary>
        /// <returns>The pivot table named style object</returns>
        public ExcelPivotTableNamedStyle PivotTableStyle
        {
            get
            {
                return _tableNamedStyle as ExcelPivotTableNamedStyle;
            }
        }
        /// <summary>
        /// Returns the table named style object as a named style that can be applied to both tables and pivot tables.
        /// </summary>
        /// <returns>The shared table named style object</returns>
        public ExcelTableAndPivotTableNamedStyle TableAndPivotTableStyle
        {
            get
            {
                return _tableNamedStyle as ExcelTableAndPivotTableNamedStyle;
            }
        }
    }
}
