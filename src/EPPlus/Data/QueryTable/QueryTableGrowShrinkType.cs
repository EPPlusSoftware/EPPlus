namespace OfficeOpenXml
{
    /// <summary>
    /// How to handle variable numbers of rows when refreshing a query table.
    /// </summary>
    public enum QueryTableGrowShrinkType
    {
        /// <summary>
        /// Insert Clear
        /// </summary>
        InsertClear, 
        /// <summary>
        /// Insert Delete
        /// </summary>
        InsertDelete, 
        /// <summary>
        /// Overwrite Clear
        /// </summary>
        OverwriteClear
    }
}