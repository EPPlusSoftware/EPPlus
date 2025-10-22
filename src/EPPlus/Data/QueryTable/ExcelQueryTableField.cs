namespace OfficeOpenXml.Data.QueryTable
{
    /// <summary>
    /// Represents a query table field
    /// </summary>
    public class ExcelQueryTableField
    {
        internal ExcelQueryTableField()
        {
            
        }
        /// <summary>
        /// A unique Id for the field
        /// </summary>
        public int Id { get; internal set; }
        /// <summary>
        /// A name for the field.
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// The index of the column in the table.
        /// </summary>
        public int TableColumnId { get; internal set; }
        /// <summary>
        /// If this field/column is currently clipped and thus not visible in the worksheet.
        /// This state might occur for example when a query table is defined near the edge of a worksheet or other object in the spreadsheet that can't be overwritten with external data.
        /// In this case some of the fields are displayed, but not all of them.
        /// </summary>
        public bool ClippedColumn { get; set; }
        /// <summary>
        /// If this column is a user-defined column or comes from the external data query. User defined columns shall be preserved during data refresh operations. Userdefined columns are only supported on query tables that are attached to table objects
        /// </summary>
        public bool DataBoundColumn { get; set; }
        /// <summary>
        /// If the formula in this field/column should be filled down on data refresh.
        /// </summary>
        public bool FillFormulaOnRefresh { get; set; }
        /// <summary>
        ///  If this column contains the row numbers for the records returned.
        /// </summary>
        public bool RowNumbers { get; set;}
    }
}