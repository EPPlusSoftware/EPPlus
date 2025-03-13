using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Represents a #SPILL! error
    /// </summary>
    public class ExcelSpillErrorValue : ExcelRichDataErrorValue
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        public ExcelSpillErrorValue(int rowOffset, int colOffset) : base(eErrorType.Spill)
        {
            SpillRowOffset = rowOffset;
            SpillColOffset = colOffset;
        }

        internal int SpillRowOffset { get; set; }
        internal int SpillColOffset { get; set; }
        internal bool IsPropagated
        {
            get;
            set;
        }
    }
}
