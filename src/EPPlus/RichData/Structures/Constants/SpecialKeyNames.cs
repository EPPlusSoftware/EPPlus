/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures.Constants
{
    /// <summary>
    /// See section 2.3.6.3 Special Keys and Key Flags in [MS-XLSX] - v20240416
    /// </summary>
    internal static class SpecialKeyNames
    {
        /// <summary>
        /// This key MUST reference a supporting property bag
        /// </summary>
        public const string Attribution = "_Attribution";
        /// <summary>
        /// This key MUST reference a supporting property bag that contains key value pairs (KVP) of strings. 
        /// The supporting property bag key in each pair contains a key that is localized. The corresponding 
        /// value contains a string representation of the key that is not locale-specific.
        /// </summary>
        public const string CanonicalPropertyNames = "_CanonicalPropertyNames";
        /// <summary>
        /// This key references a string value used to apply a Windows Information Protection (WIP) policy.
        /// </summary>
        public const string ClassificationId = "_ClassificationId";
        /// <summary>
        /// Some rich value types MAY put limitations on what type this key MUST reference. See individual 
        /// rich value type descriptions (under section 2.3.6.1) for any limitations. If this key is a part 
        /// of a rich value with an unknown rich value type (for more information see section 2.3.6.1.3.8), 
        /// this rich value MUST be preserved.
        /// </summary>
        public const string CRID = "_CRID";
        /// <summary>
        /// This key MUST reference a supporting property bag that contains information that can be used to 
        /// determine how the rich value is displayed.
        /// </summary>
        public const string Display = "_Display";
        /// <summary>
        /// This key MUST reference a supporting property bag that contains key value pairs (KVP) containing 
        /// rich value keys and supporting property bags containing associated rich value key flags, which 
        /// define behavior for the associated rich value key value pair (KVP).
        /// </summary>
        public const string Flags = "_Flags";
        /// <summary>
        /// This key MUST reference a supporting property bag that is a list of indices to a CT_RichStyle 
        /// (section 2.6.170). The supporting property bag key of the key value pair (KVP) in the supporting 
        /// property bag determines which rich value key the CT_RichStyle is associated with.
        /// </summary>
        public const string Format = "_Format";
        /// <summary>
        /// This key contains a value that describes the icon that can be used in render.
        /// </summary>
        public const string Icon = "_Icon";
        /// <summary>
        /// This key MUST reference a supporting property bag that contains information about the service provider.
        /// </summary>
        public const string Provider = "_Provider";
        /// <summary>
        /// This key SHOULD NOT exist in the file and will be removed when the file is saved. When a supporting 
        /// property bag references this rich value key, it indicates the supporting property bag references the 
        /// rich value itself, and not a key value pair (KVP) of the rich value.
        /// </summary>
        public const string Self = "_Self";
        /// <summary>
        /// This key MUST reference a supporting property bag that contains key value pairs (KVP) containing rich 
        /// value keys and supporting property bags containing strings, which define the label that can be used 
        /// to describe the associated rich value key value pair (KVP).
        /// </summary>
        public const string SubLabel = "_SubLabel";
        /// <summary>
        /// This key MUST reference a supporting property bag that contains information that MAY be used to customize 
        /// rich value visualizations.
        /// </summary>
        public const string ViewInfo = "_ViewInfo";
        /// <summary>
        /// This supporting property bag key is associated with a supporting property bag array of strings that 
        /// SHOULD be comprised of rich value keys in the associated rich value and can be used to display the 
        /// rich value key value pairs (KVP) in a different order.
        /// </summary>
        public const string Order = "^Order";

        internal static class Prefixes
        {
            /// <summary>
            /// A rich value key with this prefix MUST reference a CT_RichValueRelRelationship (section 2.6.241).
            /// </summary>
            public const string RvRel = "_rvRel";
        }
    }
}
