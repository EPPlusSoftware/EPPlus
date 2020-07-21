/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Declares how headers should be parsed before they are added to the worksheet
    /// </summary>
    public enum HeaderParsingTypes
    {
        /// <summary>
        /// Leaves the header as it is
        /// </summary>
        Preserve,
        /// <summary>
        /// Replaces any underscore characters with a space
        /// </summary>
        UnderscoreToSpace,
        /// <summary>
        /// Adds a space between camel cased words ('MyProp' => 'My Prop')
        /// </summary>
        CamelCaseToSpace,
        /// <summary>
        /// Replaces any underscore characters with a space and adds a space between camel cased words ('MyProp' => 'My Prop')
        /// </summary>
        UnderscoreAndCamelCaseToSpace
    }
}
