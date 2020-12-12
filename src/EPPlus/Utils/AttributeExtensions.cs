/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal static class AttributeExtensions
    {
        internal static bool HasPropertyOfType<T>(this MemberInfo member)
        {
            return member.GetCustomAttributes(typeof(T), false).FirstOrDefault() != null;
        }

        internal static T GetFirstAttributeOfType<T>(this MemberInfo member)
            where T : Attribute
        {
            return member.GetCustomAttributes(typeof(T), false).FirstOrDefault() as T;
        }

        internal static bool HasMemberWithPropertyOfType<T>(this Type type)
            where T : Attribute
        {
            var members = type.GetProperties();
            return members.Any(x => x.GetCustomAttributes(typeof(T), false).FirstOrDefault() != null);
        }

        internal static IEnumerable<T> FindAttributesOfType<T>(this Type type)
            where T : Attribute
        {
            var attributes = type.GetCustomAttributes(false);
            if(attributes == null || !attributes.Any())
            {
                return Enumerable.Empty<T>();
            }
            var result = new List<T>();
            foreach(var attr in attributes)
            {
                var typedAttr = attr as T;
                if(typedAttr != null)
                {
                    result.Add(typedAttr);
                }
            }
            return result;
        }

    }
}
