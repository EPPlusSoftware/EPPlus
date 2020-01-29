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
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Utilities
{
    /// <summary>
    /// Represent a function argument to validate
    /// </summary>
    /// <typeparam name="T">Type of the argument to validate</typeparam>
    public class ArgumentInfo<T>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="val">The argument to validate</param>
        public ArgumentInfo(T val)
        {
            Value = val;
        }

        /// <summary>
        /// The argument to validate
        /// </summary>
        public T Value { get; private set; }

        /// <summary>
        /// Variable name of the argument
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Sets the variable name of the argument.
        /// </summary>
        /// <param name="argName">The name</param>
        /// <returns></returns>
        public ArgumentInfo<T> Named(string argName)
        {
            Name = argName;
            return this;
        }
    }
}
