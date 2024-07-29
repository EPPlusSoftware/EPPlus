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

namespace OfficeOpenXml.Utils.TypeConversion
{
    internal class ValueWrapper
    {
        private readonly object _object;

        public ValueWrapper(object obj)
        {
            _object = obj;
        }

        public bool IsString
        {
            get
            {
                if (_object == null) return false;
                return _object is string;
            }
        }

        public bool IsEmptyString
        {
            get
            {
                if (_object == null) return false;
                return _object is string && _object.ToString().Trim() == string.Empty;
            }
        }

        public bool IsNumeric
        {
            get
            {
                if(_object == null) return false;
                return NumericTypeConversions.IsNumeric(_object.GetType());
            }
        }

        public bool IsDateTime
        {
            get
            {
                #if NET6_0_OR_GREATER
                    if (_object is DateOnly) return true;
                #endif
                return _object is DateTime;
            }
        }

        public bool IsTimeSpan
        {
            get
            {
                #if NET6_0_OR_GREATER
                   if (_object is TimeOnly) return true;
                #endif
                return _object is TimeSpan;
            }
        }

        public DateTime ToDateTime()
        {
            #if NET6_0_OR_GREATER
                if (_object is DateOnly dt) return dt.ToDateTime(TimeOnly.MinValue);
            #endif
            return (DateTime)_object;
        }

        public TimeSpan ToTimeSpan()
        {
            #if NET6_0_OR_GREATER
                if (_object is TimeOnly to) return to.ToTimeSpan();
            #endif
            return (TimeSpan)_object;
        }

        public double ToDouble()
        {
            if (IsDateTime)
            {
                return ToDateTime().ToOADate();
            }
            else if(IsTimeSpan)
            {
                return ToDateTime().ToOADate();
            }
            return Convert.ToDouble(_object);
        }

        public override string ToString()
        {
            return _object.ToString();
        }

        public object Object
        {
            get { return _object; }
        }
    }
}
