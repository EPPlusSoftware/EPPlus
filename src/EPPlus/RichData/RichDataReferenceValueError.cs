using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData
{
    /// <summary>
    /// Represents a rich value in the cell store
    /// </summary>
    public abstract class RichDataReferenceValueError : ExcelErrorValue
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="vmId">Value metadata id</param>
        /// <param name="refType">Rich Data Reference Type</param>
        internal RichDataReferenceValueError(uint vmId, RichDataReferenceTypes refType)
            : base(eErrorType.Value)
        {
            VmId = vmId;
            ReferenceType = refType;
        }

        /// <summary>
        /// Rich data Id (internal for EPPlus)
        /// </summary>
        internal uint VmId { get; private set; }

        internal bool IsValueError { get; private set; }

        /// <summary>
        /// Identifies what type of rich data being referenced
        /// </summary>
        public RichDataReferenceTypes ReferenceType { get; private set; }
    }
}
