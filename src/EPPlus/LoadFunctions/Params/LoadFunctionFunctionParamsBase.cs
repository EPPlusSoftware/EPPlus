/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using OfficeOpenXml.Table;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// Base class for parameter classes for Load functions
    /// </summary>
    public abstract class LoadFunctionFunctionParamsBase
    {
        /// <summary>
        /// If true a row with headers will be added above the data
        /// </summary>
        public bool PrintHeaders
        {
            get; set;
        }
        /// <summary>
        /// A custom name for the table, if created. 
        /// The TableName must be unique within the workbook and be a valid table name.
        /// </summary>
        public string TableName
        {
            get;
            set;
        }
        /// <summary>
        /// If set to another value than TableStyles.None the data will be added to a
        /// table with the specified style
        /// </summary>
        public TableStyles? TableStyle
        {
            get; set;
        } = null;

        /// <summary>
        /// If true, EPPlus will add the built in (default) styles for hyperlinks and apply them on any member
        /// that is of the <see cref="Uri"/> or <see cref="ExcelHyperLink"/> types. Default value is true.
        /// </summary>
        public bool UseBuiltInStylesForHyperlinks
        {
            get;
            set;
        } = true;

        /// <summary>
        /// Set if data should be transposed
        /// </summary>
        public bool Transpose { get; set; } = false;
    }
}
