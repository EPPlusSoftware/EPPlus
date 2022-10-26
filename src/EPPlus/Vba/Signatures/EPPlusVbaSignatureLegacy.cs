﻿/*************************************************************************************************
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
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.VBA.Signatures
{
    internal class EPPlusVbaSignatureLegacy : EPPlusVbaSignature
    {
        public EPPlusVbaSignatureLegacy(ZipPackagePart part) 
            : base(part, ExcelVbaSignatureType.Legacy)
        {
        }
    }
}
