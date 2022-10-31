/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// A reference to a twiddled type library
    /// </summary>
    public class ExcelVbaReferenceControl : ExcelVbaReference
    {
        /// <summary>
        /// Constructor.
        /// Sets ReferenceRecordID to 0x2F
        /// </summary>
        public ExcelVbaReferenceControl()
        {
            ReferenceRecordID = 0x2F;
        }
        /// <summary>
        /// LibIdExternal 
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// This corresponds to LibIdExtended in the documentation.
        /// </summary>
        [Obsolete("Use LibIdExtended instead of this.")]
        public string LibIdExternal
        {
            get
            {
                return LibIdExtended;
            }
            set
            {
                LibIdExtended = value;
            }
        }

        /// <summary>
        /// LibIdExtended 
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdExtended { get; set; }

        /// <summary>
        /// LibIdTwiddled
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdTwiddled { get; set; }
        /// <summary>
        /// A GUID that specifies the Automation type library the extended type library was generated from.
        /// </summary>
        public Guid OriginalTypeLib { get; set; }
        internal uint Cookie { get; set; }
    }
}