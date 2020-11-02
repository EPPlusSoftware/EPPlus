/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    /// <summary>
    /// Defines options for how to build a valid property or DataTable column name out of a string
    /// </summary>
    public enum NameParsingStrategy
    {
        /// <summary>
        /// Preserve the input string as it is
        /// </summary>
        Preserve,
        /// <summary>
        /// Replace any spaces with underscore
        /// </summary>
        SpaceToUnderscore,
        /// <summary>
        /// Remove all spaces
        /// </summary>
        RemoveSpace
    }
}
