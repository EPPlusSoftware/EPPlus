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
namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// Hash algorithm used for signing vba projects.
    /// </summary>
    public enum VbaSignatureHashAlgorithm
    {
        /// <summary>
        /// Specifies that the MD5 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        MD5 = 0,
        /// <summary>
        /// Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA1 = 1,
        /// <summary>
        /// Specifies that the SHA-256 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA256 = 2,
        /// <summary>
        /// Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA384 = 3,
        /// <summary>
        /// Specifies that the SHA-512 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA512 = 4
    }
}