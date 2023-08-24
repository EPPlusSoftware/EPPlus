/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/10/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// This class should be used to configure how arrays/ranges are treated as parameters to functions
    /// that can return a dynamic array.
    /// </summary>
    public class ArrayBehaviourConfig
    {
        /// <summary>
        /// A list of integers that specified the 0-based index of arguments that can be an array.
        /// </summary>
        public List<int> ArrayParameterIndexes { get; set; }

        /// <summary>
        /// Use this property in combination with <see cref="ArrayArgInterval"/>. A typical scenario would be that
        /// the first 3 arguments should be ignore and then every 3rd argument might be in array. In this scenario this
        /// property should be set to 3.
        /// </summary>
        public int IgnoreNumberOfArgsFromStart { get; set; }

        /// <summary>
        /// Indicates that every x-th argument can be an array.
        /// </summary>
        public int ArrayArgInterval { get; set; }

        /// <summary>
        /// Returns true if the 0-based <paramref name="ix">index</paramref>
        /// occurrs in the <see cref="ArrayParameterIndexes"/> list or if
        /// the index matches the configuration of <see cref="IgnoreNumberOfArgsFromStart"/>
        /// and <see cref="ArrayArgInterval"/>.
        /// </summary>
        /// <param name="ix"></param>
        /// <returns></returns>
        public bool CanBeArrayArg(int ix)
        {
            var startIndex = ix - IgnoreNumberOfArgsFromStart;
            if (startIndex < 0) return false;
            if (ArrayArgInterval > 0 && startIndex % ArrayArgInterval == 0) return true;
            return ArrayParameterIndexes != null && ArrayParameterIndexes.Contains(startIndex);
        }
    }
}
