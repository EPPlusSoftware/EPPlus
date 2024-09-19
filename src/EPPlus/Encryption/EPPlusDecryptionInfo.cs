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
#if(!NET35)
using OfficeOpenXml.Interfaces.SensitivityLabels;
using System.IO;

namespace OfficeOpenXml.Encryption
{
    /// <summary>
    /// Decryption information passed to the sensitivity label handler
    /// </summary>
    public class EPPlusDecryptionInfo : IPackageInfo
    {
        internal EPPlusDecryptionInfo()
        {
                
        }
        /// <summary>
        /// The package stream
        /// </summary>
        public MemoryStream PackageStream { get; set; }
        /// <summary>
        /// Protection information passed to the sensitivity label handler
        /// </summary>
        public object ProtectionInformation { get; set; }
        /// <summary>
        /// If a new sensibility label has been set, this property contains the id of the new label.
        /// </summary>
        public string ActiveLabelId { get; set; }
    }
}
#endif