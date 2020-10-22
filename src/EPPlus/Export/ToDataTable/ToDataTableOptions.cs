/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    /// <summary>
    /// This class contains options for the ToDataTable method of <see cref="ExcelRangeBase"/>.
    /// </summary>
    public class ToDataTableOptions
    {
        private const string DefaultColPrefix = "column";
        private const string DefaultDataTableName = "dataTable1";
        private ToDataTableOptions()
        {
            Mappings = new DataColumnMappingCollection();
            // Default values
            NameParsingStrategy = NameParsingStrategy.Preserve;
            PredefinedMappingsOnly = false;
            FirstRowIsColumnNames = true;
            DataTableName = DefaultDataTableName;
        }

        /// <summary>
        /// Returns an instance of ToDataTableOptions with default values set. <see cref="NameParsingStrategy"/> is set to <see cref="NameParsingStrategy.Preserve"/>, <see cref="PredefinedMappingsOnly"/> is set to false, <see cref="FirstRowIsColumnNames"/> is set to true
        /// </summary>
        public static ToDataTableOptions Default
        {
            get { return new ToDataTableOptions(); }    
        }

        /// <summary>
        /// Creates an instance of <see cref="ToDataTableOptions"/>. Use the <paramref name="configHandler"/> parameter to set the values on it.
        /// </summary>
        /// <param name="configHandler">Use this to configure the <see cref="ToDataTableOptions"/> instance in a lambda expression body.</param>
        /// <returns>The configured <see cref="ToDataTableOptions"/></returns>
        public static ToDataTableOptions Create(Action<ToDataTableOptions> configHandler)
        {
            var options = Default;
            configHandler.Invoke(options);
            return options;
        }
        /// <summary>
        /// If true, the first row of the range will be used to collect the column names of the <see cref="DataTable"/>. The column names will be set according to the <see cref="NameParsingStrategy"></see> used.
        /// </summary>
        public bool FirstRowIsColumnNames { get; set; }

        /// <summary>
        /// <see cref="NameParsingStrategy">NameParsingStrategy</see> to use when parsing the first row of the range to column names
        /// </summary>
        public NameParsingStrategy NameParsingStrategy { get; set; }

        public DataColumnMappingCollection Mappings { get; private set; }

        public bool PredefinedMappingsOnly { get; set; }

        public string ColumnNamePrefix { get; set; } = DefaultColPrefix;

        public string DataTableName { get; set; }

        public string DataTableNamespace { get; set; }
    }
}
