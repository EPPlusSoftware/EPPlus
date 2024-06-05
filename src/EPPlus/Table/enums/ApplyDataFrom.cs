
namespace OfficeOpenXml.Table
{
    /// <summary>
    /// Option for which data should overwrite the other in a sync.
    /// </summary>
    public enum ApplyDataFrom
    {
        /// <summary>
        /// Overwrite cells with column name data
        /// </summary>
        ColumnNamesToCells,
        /// <summary>
        /// Overwrite columnNames with cell data
        /// </summary>
        CellsToColumnNames
    }
}
