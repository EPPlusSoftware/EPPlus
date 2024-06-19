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

namespace OfficeOpenXml.Utils.TypeConversion
{
    /// <summary>
    /// Provides functionality for analyzing the properties of a type.
    /// </summary>
    /// <typeparam name="T">The type to analyze</typeparam>
    public class ReturnTypeWrapper<T>
    {
        private readonly Type _returnType;
        private readonly Type _underlyingType;

        /// <summary>
        /// Constructor
        /// </summary>
        public ReturnTypeWrapper()
        {
            _returnType = typeof(T);
            _underlyingType = Nullable.GetUnderlyingType(_returnType);
        }

        /// <summary>
        /// The type to analyze
        /// </summary>
        public Type Type
        {
            get
            {
                return IsNullable ? _underlyingType : _returnType;
            }
        }

        /// <summary>
        /// Returns true if the type to analyze is numeric.
        /// </summary>
        public bool IsNumeric
        {
            get
            {
                return NumericTypeConversions.IsNumeric(Type);
            }
        }

        /// <summary>
        /// Returns true if the type to analyze is nullable.
        /// </summary>
        public bool IsNullable
        {
            get
            {
                return _underlyingType != null;
            }
        }

        /// <summary>
        /// Returns true if the type to analyze equalse the <see cref="DateTime"/> type.
        /// </summary>
        public bool IsDateTime
        {
            get
            {
#if NET6_0_OR_GREATER
                return Type == typeof(DateTime) || Type == typeof(DateOnly);
#else
                return Type == typeof(DateTime);
#endif
            }
        }

        /// <summary>
        /// Returns true if the type to analyze equalse the <see cref="TimeSpan"/> type.
        /// </summary>
        public bool IsTimeSpan
        {
            get
            {
#if NET6_0_OR_GREATER
                return Type == typeof(TimeSpan) || Type == typeof(TimeOnly);
#else
                return Type == typeof(TimeSpan);
#endif
            }
        }
    }
}
