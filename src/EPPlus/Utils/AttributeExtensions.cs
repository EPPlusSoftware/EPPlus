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
        internal static bool HasAttributeOfType<T>(this MemberInfo member, bool? inherit = default(bool?))
        {
            //return member.GetCustomAttributes(typeof(T), false).FirstOrDefault() != null;

#if (NET35 || NET40)
            return member.GetCustomAttributes(typeof(T), inherit ?? false).FirstOrDefault() != null;
#else
            if (!inherit.HasValue)
            {
                return member.GetCustomAttributes(typeof(T)).FirstOrDefault() != null;
            }
            else
            {
                return member.GetCustomAttributes(typeof(T), inherit.Value).FirstOrDefault() != null;
            }
#endif
        }

        internal static T GetFirstAttributeOfType<T>(this MemberInfo member, bool? inherit = default(bool?))
            where T : Attribute
        {
#if (NET35 || NET40)
            return member.GetCustomAttributes(typeof(T), inherit ?? false).FirstOrDefault() as T;
#else
            if (!inherit.HasValue)
            {
                return member.GetCustomAttributes(typeof(T)).FirstOrDefault() as T;
            }
            else
            {
                return member.GetCustomAttributes(typeof(T), inherit.Value).FirstOrDefault() as T;
            }
#endif
        }

        internal static bool HasMemberWithAttributeOfType<T>(this Type type)
            where T : Attribute
        {
            var members = type.GetProperties();
            return members.Any(x => x.GetCustomAttributes(typeof(T), false).FirstOrDefault() != null);
        }

        internal static IEnumerable<T> FindAttributesOfType<T>(this Type type)
            where T : Attribute
        {
            var attributes = type.GetCustomAttributes(true);
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

        internal static bool IsComplexType(this Type type)
        {
            return type != typeof(string) && (type.IsClass || type.IsInterface);
        }
        /// <summary>
        /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        internal static string EncodeXMLAttribute(this string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }
        internal static string EncodeXMLElement(this string s)
        {
            return s.Replace("&", "&amp;").
                     Replace("<", "&lt;").
                     Replace(">", "&gt;");
        }

    }
}
