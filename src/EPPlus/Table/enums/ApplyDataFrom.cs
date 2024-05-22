
namespace OfficeOpenXml.Table.enums
{
    /// <summary>
    /// Option for which data should overwrite the other in a sync.
    /// </summary>
    public enum ApplyDataFrom
    {
        /// <summary>
        /// Overwrite cells with column name data
        /// </summary>
        ColumnNamesToRow,
        /// <summary>
        /// Overwrite columnNames with cell data
        /// </summary>
        RowToColumnNames
    }
}
