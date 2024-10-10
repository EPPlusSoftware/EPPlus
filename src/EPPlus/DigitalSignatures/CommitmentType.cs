using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Commitment types for signatures
    /// </summary>
    [Flags]
    public enum CommitmentType
    {
        /// <summary>
        /// None
        /// </summary>
        None = 0,
        /// <summary>
        /// Approved this document
        /// </summary>
        Approved = 1,
        /// <summary>
        /// Created this document
        /// </summary>
        Created = 2,
        /// <summary>
        /// Created and approved this document
        /// </summary>
        CreatedAndApproved = Created | Approved
    }
}
