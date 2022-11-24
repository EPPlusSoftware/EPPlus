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
using System.Reflection;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Represents the errortypes in excel
    /// </summary>
    public enum eErrorType
    {
        /// <summary>
        /// Division by zero
        /// </summary>
        Div0,
        /// <summary>
        /// Not applicable
        /// </summary>
        NA,
        /// <summary>
        /// Name error
        /// </summary>
        Name,
        /// <summary>
        /// Null error
        /// </summary>
        Null,
        /// <summary>
        /// Num error
        /// </summary>
        Num,
        /// <summary>
        /// Reference error
        /// </summary>
        Ref,
        /// <summary>
        /// Value error
        /// </summary>
        Value
    }

    /// <summary>
    /// Represents an Excel error.
    /// </summary>
    /// <seealso cref="eErrorType"/>
    public class ExcelErrorValue
    {
        /// <summary>
        /// Handles the convertion between <see cref="eErrorType"/> and the string values
        /// used by Excel.
        /// </summary>
        public static class Values
        {
            /// <summary>
            /// A constant for Div/0 error in Excel
            /// </summary>
            public const string Div0 = "#DIV/0!";
            /// <summary>
            /// A constant for the N/A error in Excel
            /// </summary>
            public const string NA = "#N/A";
            /// <summary>
            /// A constant for the Name error in Excel
            /// </summary>
            public const string Name = "#NAME?";
            /// <summary>
            /// A constant for the Numm error in Excel
            /// </summary>
            public const string Null = "#NULL!";
            /// <summary>
            /// A constant for the Num error in Excel
            /// </summary>
            public const string Num = "#NUM!";
            /// <summary>
            /// A constant for the Ref error in Excel
            /// </summary>
            public const string Ref = "#REF!";
            /// <summary>
            /// A constant for the Value error in Excel
            /// </summary>
            public const string Value = "#VALUE!";

            private static Dictionary<string, eErrorType> _values = new Dictionary<string, eErrorType>()
                {
                    {Div0, eErrorType.Div0},
                    {NA, eErrorType.NA},
                    {Name, eErrorType.Name},
                    {Null, eErrorType.Null},
                    {Num, eErrorType.Num},
                    {Ref, eErrorType.Ref},
                    {Value, eErrorType.Value}
                };

            /// <summary>
            /// Returns true if the supplied <paramref name="candidate"/> is an excel error.
            /// </summary>
            /// <param name="candidate"></param>
            /// <returns></returns>
            public static bool IsErrorValue(object candidate)
            {
                if(candidate == null || !(candidate is ExcelErrorValue)) return false;
                var candidateString = candidate.ToString();
                return (!string.IsNullOrEmpty(candidateString) && _values.ContainsKey(candidateString));
            }

            /// <summary>
            /// Returns true if the supplied <paramref name="candidate"/> is an excel error.
            /// </summary>
            /// <param name="candidate"></param>
            /// <returns></returns>
            public static bool StringIsErrorValue(string candidate)
            {
                return (!string.IsNullOrEmpty(candidate) && _values.ContainsKey(candidate));
            }

            /// <summary>
            /// Converts a string to an <see cref="eErrorType"/>
            /// </summary>
            /// <param name="val"></param>
            /// <returns></returns>
            /// <exception cref="ArgumentException">Thrown if the supplied value is not an Excel error</exception>
            public static eErrorType ToErrorType(string val)
            {
                if (string.IsNullOrEmpty(val) || !_values.ContainsKey(val))
                {
                    throw new ArgumentException("Invalid error code " + (val ?? "<empty>"));
                }
                return _values[val];
            }
        }

        /// <summary>
        /// Creates an <see cref="ExcelErrorValue"/> from a <see cref="ExcelErrorValue"/>
        /// </summary>
        /// <param name="errorType">The type of error to create</param>
        /// <returns>The <see cref="ExcelErrorValue"/></returns>
        public static ExcelErrorValue Create(eErrorType errorType)
        {
            return new ExcelErrorValue(errorType);
        }

        /// <summary>
        /// Parses a error value string and returns the <see cref="ExcelErrorValue"/>
        /// </summary>
        /// <param name="val">The error code</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">Is thrown when <paramref name="val"/> is empty</exception>
        /// <exception cref="ArgumentException">Is thrown when <paramref name="val"/> is not a valid Excel error.</exception>
        /// <exception cref="ArgumentException">If the argument cannot be converted.</exception>
        public static ExcelErrorValue Parse(string val)
        {
            if (Values.StringIsErrorValue(val))
            {
                return new ExcelErrorValue(Values.ToErrorType(val));
            }
            if(string.IsNullOrEmpty(val)) throw new ArgumentNullException("val");
            throw new ArgumentException("Not a valid error value: " + val);
        }

        internal static bool IsErrorValue(string val)
        {
            return Values.StringIsErrorValue(val);
        }

        private ExcelErrorValue(eErrorType type)
        {
            Type=type; 
        }

        /// <summary>
        /// The error type
        /// </summary>
        public eErrorType Type { get; private set; }

        /// <summary>
        /// Returns the string representation of the error type
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            switch(Type)
            {
                case eErrorType.Div0:
                    return Values.Div0;
                case eErrorType.NA:
                    return Values.NA;
                case eErrorType.Name:
                    return Values.Name;
                case eErrorType.Null:
                    return Values.Null;
                case eErrorType.Num:
                    return Values.Num;
                case eErrorType.Ref:
                    return Values.Ref;
                case eErrorType.Value:
                    return Values.Value;
                default:
                    throw(new ArgumentException("Invalid errortype"));
            }
        }
        /// <summary>
        /// Operator for addition.
        /// </summary>
        /// <param name="v1">Left side</param>
        /// <param name="v2">Right side</param>
        /// <returns>Return the error value in V2</returns>
        public static ExcelErrorValue operator +(object v1, ExcelErrorValue v2)
        {
            return v2;
        }
        /// <summary>
        /// Operator for addition.
        /// </summary>
        /// <param name="v1">Left side</param>
        /// <param name="v2">Right side</param>
        /// <returns>Return the error value in V1</returns>
        public static ExcelErrorValue operator +(ExcelErrorValue v1, ExcelErrorValue v2)
        {
            return v1;
        }

        /// <summary>
        /// Calculates a hash code for the object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        /// <summary>
        /// Checks if the object is equals to another
        /// </summary>
        /// <param name="obj">The object to compare</param>
        /// <returns>True if equals</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ExcelErrorValue)) return false;
            return ((ExcelErrorValue) obj).ToString() == this.ToString();
        }
    }
}
