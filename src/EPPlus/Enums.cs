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
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Tells how cells should be shifted in a delete operation
    /// </summary>
    public enum eShiftTypeDelete
    {
        /// <summary>
        /// Cells in the range are shifted to the left
        /// </summary>
        Left,
        /// <summary>
        /// Cells in the range are shifted upwards
        /// </summary>
        Up,
        /// <summary>
        /// The range for the entire row is used in the shift operation
        /// </summary>
        EntireRow,
        /// <summary>
        /// The range for the entire column is used in the shift operation
        /// </summary>
        EntireColumn
    }
    /// <summary>
    /// Tells how cells should be shifted in a insert operation
    /// </summary>
    public enum eShiftTypeInsert
    {
        /// <summary>
        /// Cells in the range are shifted to the right
        /// </summary>
        Right,
        /// <summary>
        /// Cells in the range are shifted downwards
        /// </summary>
        Down,
        /// <summary>   
        /// The range for the entire row is used in the shift operation
        /// </summary>
        EntireRow,
        /// <summary>
        /// The range for the entire column is used in the shift operation
        /// </summary>
        EntireColumn
    }
    /// <summary>
    /// Algorithm for password hash
    /// </summary>
    internal enum eProtectedRangeAlgorithm
    {
        /// <summary>
        /// Specifies that the MD2 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        MD2,
        /// <summary>
        /// Specifies that the MD4 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        MD4,
        /// <summary>
        /// Specifies that the MD5 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        MD5,
        /// <summary>
        /// Specifies that the RIPEMD-128 algorithm, as defined by RFC 1319, shall be used.
        /// </summary>
        RIPEMD128,
        /// <summary>
        /// Specifies that the RIPEMD-160 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        RIPEMD160,
        /// <summary>
        /// Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA1,
        /// <summary>
        /// Specifies that the SHA-256 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA256,
        /// <summary>
        /// Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        SHA384,
        /// <summary>
        /// Specifies that the SHA-512 algorithm, as defined by ISO/IEC10118-3:2004 shall be used.
        /// </summary>
        SHA512,
        /// <summary>
        /// Specifies that the WHIRLPOOL algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
        /// </summary>
        WHIRLPOOL
    }
    /// <summary>
    /// Maps to DotNetZips CompressionLevel enum
    /// </summary>
    public enum CompressionLevel
    {
        /// <summary>
        /// Level 0, no compression
        /// </summary>
        Level0 = 0,
        /// <summary>
        /// No compression
        /// </summary>
        None = 0,
        /// <summary>
        /// Level 1, Best speen
        /// </summary>
        Level1 = 1,
        /// <summary>
        /// 
        /// </summary>
        BestSpeed = 1,
        /// <summary>
        /// Level 2
        /// </summary>
        Level2 = 2,
        /// <summary>
        /// Level 3
        /// </summary>
        Level3 = 3,
        /// <summary>
        /// Level 4
        /// </summary>
        Level4 = 4,
        /// <summary>
        /// Level 5
        /// </summary>
        Level5 = 5,
        /// <summary>
        /// Level 6
        /// </summary>
        Level6 = 6,
        /// <summary>
        /// Default, Level 6
        /// </summary>
        Default = 6,
        /// <summary>
        /// Level 7
        /// </summary>
        Level7 = 7,
        /// <summary>
        /// Level 8
        /// </summary>
        Level8 = 8,
        /// <summary>
        /// Level 9
        /// </summary>
        BestCompression = 9,
        /// <summary>
        /// Best compression, Level 9
        /// </summary>
        Level9 = 9,
    }
    /// <summary>
    /// Specifies with license EPPlus is used under.
    /// Licensetype must be specified in order to use the library
    /// <seealso cref="ExcelPackage.LicenseContext"/>
    /// </summary>
    public enum LicenseContext
    {
        /// <summary>
        /// You comply with the Polyform Non Commercial License.
        /// See https://polyformproject.org/licenses/noncommercial/1.0.0/
        /// </summary>
        NonCommercial = 0,
        /// <summary>
        /// You have a commercial license purchased at https://epplussoftware.com/licenseoverview
        /// </summary>
        Commercial = 0
    }
}
