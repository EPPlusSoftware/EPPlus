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
using System.Reflection;
namespace OfficeOpenXml.Compatibility
{
    internal class TypeCompat
    {
        public static bool IsPrimitive(object v)
        {
#if (Core)            
            return v.GetType().GetTypeInfo().IsPrimitive;
#else
            return v.GetType().IsPrimitive;
#endif

        }
        public static bool IsSubclassOf(Type t, Type c)
        {
#if (Core)            
            return t.GetTypeInfo().IsSubclassOf(c);
#else
            return t.IsSubclassOf(c);
#endif
        }

        internal static bool IsGenericType(Type t)
        {
#if (Core)            
            return t.GetTypeInfo().IsGenericType;
#else
            return t.IsGenericType;
#endif

        }
        public static object GetPropertyValue(object v, string name)
        {
#if (Core)
            return v.GetType().GetTypeInfo().GetProperty(name).GetValue(v, null);
#else
            return v.GetType().GetProperty(name).GetValue(v, null);
#endif
        }
    }
}
