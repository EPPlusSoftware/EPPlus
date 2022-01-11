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
    internal static class NumericTypeConversions
    {
        private static readonly Dictionary<Type, Func<object, object>> _numericTypes = new Dictionary<Type, Func<object, object>>
        {
            { typeof(byte), (o) => Convert.ToByte(o) },
            { typeof(uint), (o) => Convert.ToUInt32(o) },
            { typeof(int), (o) => Convert.ToInt32(o) },
            { typeof(float), (o) => {
                if(o == null) return null;
                float output;
                if(float.TryParse(o.ToString(), out output)) return output;
                return null;
            } },
            { typeof(double), (o) => Convert.ToDouble(o) },
            { typeof(decimal), (o) => Convert.ToDecimal(o) },
            { typeof(ulong), (o) => Convert.ToUInt64(o) },
            { typeof(long), (o) => Convert.ToInt64(o) },
            { typeof(ushort), (o) => Convert.ToUInt16(o) },
            { typeof(short), (o) => Convert.ToInt16(o) }
        };

        public static bool IsNumeric(Type type)
        {
            return _numericTypes.ContainsKey(type);
        }

        public static bool TryConvert(object obj, out object convertedObj, Type convertToType)
        {
            convertedObj = obj;
            try
            {
                if (_numericTypes.ContainsKey(convertToType))
                {
                    var conversionFunc = _numericTypes[convertToType];
                    convertedObj = conversionFunc(obj);
                    return true;
                }
                return false;
            }
            catch 
            {
                return false;
            }
        }
    }
}
