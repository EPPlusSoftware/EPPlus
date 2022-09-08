/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/05/2022         EPPlus Software AB       EPPlus 6.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.VBA.ContentHash
{
    /// <summary>
    /// Hash algorithms for usage when signing VBA
    /// </summary>
    internal enum ExcelVbaSignatureHashAlgorithm
    {
        /// <summary>
        /// MD5 hash algorithm
        /// </summary>
        MD5,
        /// <summary>
        /// SHA1 hash algorithm
        /// </summary>
        SHA1,
        /// <summary>
        /// SHA256 hash algorithm
        /// </summary>
        SHA256,
        /// <summary>
        /// SHA384 hash algorithm
        /// </summary>
        SHA384,
        /// <summary>
        /// SHA512 hash algorithm
        /// </summary>
        SHA512
    }
}
