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
    internal class TypeConvertUtil<TReturnType>
    {
        internal TypeConvertUtil(object o)
        {
            Value = new ValueWrapper(o);
            ReturnType = new ReturnTypeWrapper<TReturnType>();
        }

        public ReturnTypeWrapper<TReturnType> ReturnType
        {
            get;
            private set;
        }

        public ValueWrapper Value
        {
            get;
            private set;
        }

        public object ConvertToReturnType()
        {
            if (ReturnType.IsNullable && Value.IsEmptyString)
            {
                return null;
            }
            if (NumericTypeConversions.IsNumeric(ReturnType.Type))
            {
                object convertedObj;      
                if(NumericTypeConversions.TryConvert(Value.Object, out convertedObj, ReturnType.Type))
                {
                    return convertedObj;
                }
                return default(TReturnType);
            }
            return Value.Object;
        }

        public bool TryGetDateTime(out object returnDate)
        {
            returnDate = default;
            if (!ReturnType.IsDateTime) return false;
            if (Value.Object is double)
            {
                returnDate = DateTime.FromOADate(Value.ToDouble());
                return true;
            }
            if (Value.IsTimeSpan)
            {
                returnDate = new DateTime(Value.ToTimeSpan().Ticks);
                return true;
            }
            if (Value.IsString)
            {
                if (DateTime.TryParse(Value.ToString(), out DateTime dt))
                {
                    returnDate = dt;
                    return true;
                }
            }
            return false;
        }

        public bool TryGetTimeSpan(out object timeSpan)
        {
            timeSpan = default;
            if (!ReturnType.IsTimeSpan) return false;
            if (Value.Object is long)
            {
                timeSpan = new TimeSpan(Convert.ToInt64(Value.Object));
                return true;
            }
            if(Value.Object is double)
            {
                timeSpan = new TimeSpan(DateTime.FromOADate((double)Value.Object).Ticks);
                return true;
            }
            if (Value.IsDateTime)
            {
                timeSpan = new TimeSpan(Value.ToDateTime().Ticks);
                return true;
            }
            if (Value.IsString)
            {
                TimeSpan ts;
                if (TimeSpan.TryParse(Value.ToString(), out ts))
                {
                    timeSpan = ts;
                    return true;
                }
                throw new FormatException(Value.ToString() + " could not be parsed to a TimeSpan");
            }
            return false;
        }
    }
}
