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
        private List<string> _primaryKeyFields = new List<string>();
        private List<int> _primaryKeyIndexes = new List<int>();
        private ToDataTableOptions()
        {
            Mappings = new DataColumnMappingCollection();
            // Default values
            ColumnNameParsingStrategy = NameParsingStrategy.Preserve;
            ExcelErrorParsingStrategy = ExcelErrorParsingStrategy.HandleExcelErrorsAsBlankCells;
            PredefinedMappingsOnly = false;
            FirstRowIsColumnNames = true;
            DataTableName = DefaultDataTableName;
        }

        internal IEnumerable<string> PrimaryKeyNames
        {
            get { return _primaryKeyFields; }
        }

        internal IEnumerable<int> PrimaryKeyIndexes
        {
            get { return _primaryKeyIndexes; }
        }

        /// <summary>
        /// Returns an instance of ToDataTableOptions with default values set. <see cref="ColumnNameParsingStrategy"/> is set to <see cref="NameParsingStrategy.Preserve"/>, <see cref="PredefinedMappingsOnly"/> is set to false, <see cref="FirstRowIsColumnNames"/> is set to true
        /// </summary>
        internal static ToDataTableOptions Default
        {
            get { return new ToDataTableOptions(); }    
        }

        /// <summary>
        /// Creates an instance of ToDataTableOptions with default values set.
        /// </summary>
        /// <returns></returns>
        /// <seealso cref="Default"/>
        public static ToDataTableOptions Create()
        {
            return new ToDataTableOptions();
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
        /// If true, the first row of the range will be used to collect the column names of the <see cref="DataTable"/>. The column names will be set according to the <see cref="ColumnNameParsingStrategy"></see> used.
        /// </summary>
        public bool FirstRowIsColumnNames { get; set; }

        /// <summary>
        /// <see cref="ColumnNameParsingStrategy">NameParsingStrategy</see> to use when parsing the first row of the range to column names
        /// </summary>
        public NameParsingStrategy ColumnNameParsingStrategy { get; set; }

        /// <summary>
        /// Number of rows that will be skipped from the start (top) of the range. If <see cref="FirstRowIsColumnNames"/> is true, this will be applied after the first row (column names) has been read.
        /// </summary>
        public int SkipNumberOfRowsStart { get; set; }

        /// <summary>
        /// Number of rows that will be skipped from the end (bottom) of the range.
        /// </summary>
        public int SkipNumberOfRowsEnd { get; set; }

        /// <summary>
        /// Sets how Excel error values are handled when detected.
        /// </summary>
        public ExcelErrorParsingStrategy ExcelErrorParsingStrategy { get; set; }

        /// <summary>
        /// Sets how empty rows in the range are handled when detected
        /// </summary>
        public EmptyRowsStrategy EmptyRowStrategy { get; set; }

        /// <summary>
        /// Mappings that specifies columns from the range and how these should be mapped to the <see cref="DataTable"/>
        /// </summary>
        /// <seealso cref="DataColumnMapping"/>
        public DataColumnMappingCollection Mappings { get; private set; }

        /// <summary>
        /// If true, only columns that are specified in the <see cref="Mappings"></see> collection are included in the DataTable.
        /// </summary>
        public bool PredefinedMappingsOnly { get; set; }

        /// <summary>
        /// If no column names are specified, this prefix will be used followed by a number
        /// </summary>
        public string ColumnNamePrefix { get; set; } = DefaultColPrefix;

        /// <summary>
        /// Name of the data table
        /// </summary>
        public string DataTableName { get; set; }

        /// <summary>
        /// Namespace of the data table
        /// </summary>
        public string DataTableNamespace { get; set; }

        /// <summary>
        /// If true, the <see cref="DataColumnMapping.AllowNull"/> will be overridden and
        /// null values will be allowed in all columns.
        /// </summary>
        public bool AlwaysAllowNull { get; set; }

        /// <summary>
        /// Set to true if the worksheet is contains transposed data.
        /// </summary>
        public bool DataIsTransposed { get; set; }

        /// <summary>
        /// Sets the primary key of the data table. 
        /// </summary>
        /// <param name="columnNames">The name or names of one or more column in the <see cref="System.Data.DataTable"/> that constitutes the primary key</param>
        public void SetPrimaryKey(params string[] columnNames)
        {
            _primaryKeyFields.Clear();
            _primaryKeyFields.AddRange(columnNames);
            _primaryKeyIndexes.Clear();
        }

        /// <summary>
        /// Sets the primary key of the data table. 
        /// </summary>
        /// <param name="zeroBasedRangeIndexes">The index or indexes of one or more column in the range that builds up the primary key of the <see cref="System.Data.DataTable"/></param>
        public void SetPrimaryKey(params int[] zeroBasedRangeIndexes)
        {
            _primaryKeyIndexes.Clear();
            _primaryKeyIndexes.AddRange(zeroBasedRangeIndexes);
            _primaryKeyFields.Clear();
        }
    }
}
