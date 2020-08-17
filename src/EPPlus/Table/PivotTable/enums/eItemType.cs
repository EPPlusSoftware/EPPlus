using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// The item type for a pivot table field
    /// </summary>
    public enum eItemType
    {
        /// <summary>
        /// The pivot item represents data.
        /// </summary>
        Data,
        /// <summary>
        /// The pivot item represents an "average" aggregate function.
        /// </summary>
        Avg,
        /// <summary>
        /// The pivot item represents a blank line.
        /// </summary>
        Blank,
        /// <summary>
        /// The pivot item represents custom the "count" aggregate function.
        /// </summary>
        Count,
        /// <summary>
        /// The pivot item represents custom the "count numbers" aggregate.
        /// </summary>
        CountA,
        /// <summary>
        /// The pivot item represents the default type for this PivotTable. 
        /// The default pivot item type is the "total" aggregate function.
        /// </summary>
        Default,
        /// <summary>
        /// The pivot items represents the grand total line.
        /// </summary>
        Grand,
        /// <summary>
        /// The pivot item represents the "maximum" aggregate function.
        /// </summary>
        Max,
        /// <summary>
        /// The pivot item represents the "minimum" aggregate function.
        /// </summary>
        Min,
        /// <summary>
        /// The pivot item represents the "product" function.
        /// </summary>
        Product,
        /// <summary>
        /// The pivot item represents the "standard deviation" aggregate function.
        /// </summary>
        StdDev,
        /// <summary>
        /// The pivot item represents the "standard deviation population" aggregate function.
        /// </summary>
        StdDevP,
        /// <summary>
        /// The pivot item represents the "sum" aggregate value.
        /// </summary>
        Sum,
        /// <summary>
        /// The pivot item represents the "variance" aggregate value.
        /// </summary>
        Var,
        /// <summary>
        /// The pivot item represents the "variance population" aggregate value.
        /// </summary>
        VarP
    }
}
