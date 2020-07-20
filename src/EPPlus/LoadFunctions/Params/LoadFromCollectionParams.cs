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
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Parameters for the LoadFromCollection method
    /// </summary>
    public class LoadFromCollectionParams : LoadFunctionFunctionParamsBase
    {
        /// <summary>
        /// Default value for the BindingFlags property
        /// </summary>
        public const BindingFlags DefaultBindingFlags = BindingFlags.Public | BindingFlags.Instance;

        /// <summary>
        /// The <see cref="BindingFlags"/> used when reading properties via reflection.
        /// </summary>
        public BindingFlags BindingFlags { get; set; } = DefaultBindingFlags;

        /// <summary>
        /// If not null, this specifies the members that should be used. Any member not present will be ignored.
        /// </summary>
        public MemberInfo[] Members { get; set; }

        /// <summary>
        /// Sets how headers should be parsed before added to the worksheet, see <see cref="HeaderParsingTypes"/>
        /// </summary>
        public HeaderParsingTypes HeaderParsingType { get; set; } = HeaderParsingTypes.UnderscoreToSpace;
    }
}
