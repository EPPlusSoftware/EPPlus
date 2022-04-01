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

namespace OfficeOpenXml.FormulaParsing.Exceptions
{
    /// <summary>
    /// Represents an Excel error code.
    /// </summary>
    public class ExcelErrorCodes
    {
        private ExcelErrorCodes(string code)
        {
            Code = code;
        }
        /// <summary>
        /// The error code
        /// </summary>
        public string Code
        {
            get;
            private set;
        }

        /// <summary>
        /// Returns the hash code for this string.
        /// </summary>
        /// <returns>The hash code</returns>
        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }
        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="obj">The object to compare with the current object.</param>
        /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (obj is ExcelErrorCodes)
            {
                return ((ExcelErrorCodes)obj).Code.Equals(Code);
            }
 	        return false;
        }

        /// <summary>
        /// Equal operator
        /// </summary>
        /// <param name="c1">The first error code to match</param>
        /// <param name="c2">The second error code to match</param>
        /// <returns></returns>
        public static bool operator == (ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return c1.Code.Equals(c2.Code);
        }
        /// <summary>
        /// Not equal operator
        /// </summary>
        /// <param name="c1">The first error code to match</param>
        /// <param name="c2">The second error code to match</param>
        /// <returns></returns>
        public static bool operator !=(ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return !c1.Code.Equals(c2.Code);
        }

        private static readonly IEnumerable<string> Codes = new List<string> { Value.Code, Name.Code, NoValueAvaliable.Code };

        /// <summary>
        /// Returns true if <paramref name="valueToTest"/> matches an error code.
        /// </summary>
        /// <param name="valueToTest"></param>
        /// <returns></returns>
        public static bool IsErrorCode(object valueToTest)
        {
            if (valueToTest == null)
            {
                return false;
            }
            var candidate = valueToTest.ToString();
            if (Codes.FirstOrDefault(x => x == candidate) != null)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Represents a cell value error 
        /// </summary>
        public static ExcelErrorCodes Value
        {
            get { return new ExcelErrorCodes("#VALUE!"); }
        }

        /// <summary>
        /// Represents a cell name error 
        /// </summary>
        public static ExcelErrorCodes Name
        {
            get { return new ExcelErrorCodes("#NAME?"); }
        }
        /// <summary>
        /// Reprecents a N/A error
        /// </summary>
        public static ExcelErrorCodes NoValueAvaliable
        {
            get { return new ExcelErrorCodes("#N/A"); }
        }
    }
}
