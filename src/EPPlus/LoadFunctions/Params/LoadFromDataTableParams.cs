using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Parameters from the <see cref="ExcelRangeBase.LoadFromDataTable(System.Data.DataTable, Action{LoadFromDataTableParams})"/>
    /// </summary>
    public class LoadFromDataTableParams
    {
        /// <summary>
        /// If the Caption of the DataColumn should be used as header.
        /// </summary>
        public bool PrintHeaders { get; set; }

        /// <summary>
        /// The table style to use on the table created for the imported data.
        /// null means that no table is created.
        /// </summary>
        public TableStyles? TableStyle { get; set; }

        /// <summary>
        /// Transpose the loaded data
        /// </summary>
        public bool Transpose { get; set; } = false;
    }
}
