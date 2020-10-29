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
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    public class DataColumnMapping
    {
        internal DataColumnMapping(DataColumn dataColumn)
        {
            Require.That(dataColumn).IsNotNull();
            Require.That(dataColumn.ColumnName).IsNotNullOrEmpty();
            Require.That(dataColumn.DataType).IsNotNull();
            DataColumn = dataColumn;
        }

        internal DataColumnMapping()
        {

        }

        internal bool HasDataColumn => this.DataColumn != null;

        /// <summary>
        /// The <see cref="System.Data.DataColumn"/> used for the mapping
        /// </summary>
        public DataColumn DataColumn
        {
            get;
            private set;
        }

        /// <summary>
        /// Zero based index of the mappings column in the range
        /// </summary>
        public int ZeroBasedColumnIndexInRange { get; set; }

        private string _dataColumnName;

        /// <summary>
        /// Name of the data column, corresponds to <see cref="System.Data.DataColumn.ColumnName"/>
        /// </summary>
        public string DataColumnName
        {
            get { return HasDataColumn ? DataColumn.ColumnName : _dataColumnName; }
            set
            {
                if(HasDataColumn)
                {
                    DataColumn.ColumnName = value;
                }
                else
                {
                    _dataColumnName = value;
                }
            }
        }

        private Type _dataColumnType;
        /// <summary>
        /// <see cref="Type">Type</see> of the column, corresponds to <see cref="System.Data.DataColumn.DataType"/>
        /// </summary>
        public Type ColumnDataType
        {
            get
            {
                if(HasDataColumn)
                {
                    return DataColumn.DataType;
                }
                else
                {
                    return _dataColumnType;
                }
            }
            set
            {
                if(HasDataColumn)
                {
                    DataColumn.DataType = value;
                }
                else
                {
                    _dataColumnType = value;
                }
            }
        }

        private bool _allowNull;
        /// <summary>
        /// Indicates whether empty cell values should be allowed. Corresponds to <see cref="System.Data.DataColumn.AllowDBNull"/>
        /// </summary>
        public bool AllowNull
        {
            get
            {
                if(HasDataColumn)
                {
                    return DataColumn.AllowDBNull;
                }
                else
                {
                    return _allowNull;
                }
            }
            set
            {
                if(HasDataColumn)
                {
                    DataColumn.AllowDBNull = value;
                }
                else
                {
                    _allowNull = value;
                }
            }
        }

        /// <summary>
        /// A function which allows 
        /// </summary>
        public Func<object, object> TransformCellValue
        {
            get; set;
        }

        internal void Validate()
        {
            if(string.IsNullOrEmpty(DataColumnName)) throw new ArgumentNullException("DataColumnName");
            if (ZeroBasedColumnIndexInRange < 0) throw new ArgumentOutOfRangeException("ZeroBasedColumnIndex");
        }
    }
}
