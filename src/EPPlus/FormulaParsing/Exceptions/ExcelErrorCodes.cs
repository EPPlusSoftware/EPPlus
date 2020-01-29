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
    public class ExcelErrorCodes
    {
        private ExcelErrorCodes(string code)
        {
            Code = code;
        }

        public string Code
        {
            get;
            private set;
        }

        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }

        public override bool  Equals(object obj)
        {
            if (obj is ExcelErrorCodes)
            {
                return ((ExcelErrorCodes)obj).Code.Equals(Code);
            }
 	        return false;
        }

        public static bool operator == (ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return c1.Code.Equals(c2.Code);
        }

        public static bool operator !=(ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return !c1.Code.Equals(c2.Code);
        }

        private static readonly IEnumerable<string> Codes = new List<string> { Value.Code, Name.Code, NoValueAvaliable.Code };

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

        public static ExcelErrorCodes Value
        {
            get { return new ExcelErrorCodes("#VALUE!"); }
        }

        public static ExcelErrorCodes Name
        {
            get { return new ExcelErrorCodes("#NAME?"); }
        }

        public static ExcelErrorCodes NoValueAvaliable
        {
            get { return new ExcelErrorCodes("#N/A"); }
        }
    }
}
