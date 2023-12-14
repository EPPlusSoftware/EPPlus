/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/21/2023         EPPlus Software AB           EPPlus 7.0.2
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    internal static class TypeExtensions
    {
        public static Type GetTypeOrUnderlyingType(this Type type)
        {
            var t = type;
            var ut = Nullable.GetUnderlyingType(t);
            if (ut != null)
            {
                t = ut;
            }
            return t;
        }

        public static bool IsComplexType(this Type type)
        {
            return type != typeof(string) && (type.IsClass || type.IsInterface || type.IsGenericType);
        }

        private static bool ListContainsMember(List<MemberInfo> members, MemberInfo member)
        {
            return members.Any(x => 
                x.Name == member.Name 
                && (x.DeclaringType == member.DeclaringType
                ||
                member.DeclaringType.IsSubclassOf(x.DeclaringType)));
        }

        public static IEnumerable<MemberInfo> GetLoadFromCollectionMembers(this Type type, BindingFlags bindingFlags, MemberInfo[] filterMembers)
        {
            IEnumerable<MemberInfo> members = type.GetProperties(bindingFlags).Cast<MemberInfo>().Where(x => x.ShouldBeIncluded());
            var membersList = members.ToList();
            if (filterMembers == null) return membersList;
            var hs = new HashSet<MemberInfo>(membersList);
            foreach (var filterMember in filterMembers)
            {
                var fmt = filterMember.DeclaringType;
                if (fmt.IsSubclassOf(type))
                {
                    membersList.Add(filterMember);
                }
                else if(filterMember.DeclaringType == type && !ListContainsMember(membersList, filterMember))
                {
                    membersList.Add(filterMember);
                }
            }
            return membersList;
        }
    }
}
