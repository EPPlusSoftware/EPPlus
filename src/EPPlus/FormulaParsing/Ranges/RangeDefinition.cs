using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// Represents the size of a range
    /// </summary>
    public struct RangeDefinition
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="nRows">Number of rows</param>
        /// <param name="nCols">Number of columns</param>
        public RangeDefinition(int nRows, short nCols)
        {
            NumberOfCols = nCols;
            NumberOfRows = nRows;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fromCol">From column</param>
        /// <param name="fromRow">From row</param>
        /// <param name="toCol">To column</param>
        /// <param name="toRow">To row</param>
        public RangeDefinition(short fromCol, int fromRow, short toCol, int toRow)
        {
            NumberOfCols = (short)(toCol - fromCol);
            NumberOfRows = toRow - fromRow;
        }

        /// <summary>
        /// Number of columns in the range
        /// </summary>
        public short NumberOfCols { get; private set; }

        /// <summary>
        /// Number of rows in the range
        /// </summary>
        public int NumberOfRows { get; private set; }
    }
}
