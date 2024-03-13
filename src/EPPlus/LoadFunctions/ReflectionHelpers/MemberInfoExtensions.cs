/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/7/2023         EPPlus Software AB       EPPlus 7.0.4
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    internal static class MemberInfoExtensions
    {
        public static bool ShouldBeIncluded(this MemberInfo memberInfo)
        {
            if (memberInfo.HasAttributeOfType<EpplusIgnore>()) return false;
            if(memberInfo.MemberType == MemberTypes.Method)
            {
                if (memberInfo.Name.StartsWith("get_") || memberInfo.Name.StartsWith("set_"))
                    return false;
                else if (memberInfo.DeclaringType == typeof(object))
                    return false;
            }
            return memberInfo.MemberType == MemberTypes.Field ||
                memberInfo.MemberType == MemberTypes.Property ||
                memberInfo.MemberType == MemberTypes.Method;
        }

        public static Type GetMemberType(this MemberInfo memberInfo)
        {
            switch(memberInfo.MemberType)
            {
#if !NET35
                case MemberTypes.TypeInfo:
                    return ((TypeInfo)memberInfo).GetType();
#endif
                case MemberTypes.Custom:
                    if(memberInfo is DictionaryItemMemberInfo dimi)
                    {
                        return typeof(Dictionary<string, object>);
                    }
                    throw new InvalidOperationException($"Member {memberInfo.Name} is not a Field, Property or Method");
                case MemberTypes.Field:
                    return ((FieldInfo)memberInfo).FieldType;
                case MemberTypes.Property:
                    return ((PropertyInfo)memberInfo).PropertyType;
                case MemberTypes.Method:
                    return ((MethodInfo)memberInfo).ReturnType;
                default:
                    throw new InvalidOperationException($"Member {memberInfo.Name} is not a Field, Property or Method");
            }
        }

        public static object GetValue(this MemberInfo memberInfo, object obj, BindingFlags bindingFlags)
        {
            if(obj == null)
            {
                return null;
            }
            var ot = obj.GetType();
            var mt = memberInfo.DeclaringType;
            if(ot != mt && obj.GetType().GetMember(memberInfo.Name, bindingFlags).Length == 0)
            {
                return null;
            }
            object retVal;
            if (memberInfo is PropertyInfo pi)
            {
                retVal = pi.GetValue(obj, null);
            }
            else if (memberInfo is FieldInfo fi)
            {
                retVal = fi.GetValue(obj);
            }
            else if (memberInfo is MethodInfo mi)
            {
                retVal = mi.Invoke(obj, null);
            }
            else if(memberInfo is DictionaryItemMemberInfo dim)
            {
                retVal = dim.GetValue(obj);
            }
            else
            {
                throw new NotSupportedException("Invalid member: '" + memberInfo.Name + "', not supported member type '" + memberInfo.GetType().FullName + "'");
            }
            return retVal;
        }
    }
}
