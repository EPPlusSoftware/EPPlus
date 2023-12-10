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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal static class MemberHelper
    {
        internal static void ValidateType(MemberInfo member, HashSet<Type> includedTypes)
        {
            var isValid = false;
            foreach (var includedType in includedTypes)
            {

                if (member.DeclaringType == includedType
                    || member.DeclaringType.IsAssignableFrom(includedType)
                    || member.DeclaringType.IsSubclassOf(includedType))
                {
                    isValid = true;
                    break;
                }
            }
            if (!isValid) throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
        }

        internal static Type GetTypeByMemberInfo(MemberInfo member)
        {
            switch (member.MemberType)
            {
                case MemberTypes.Field:
                    return ((FieldInfo)member).FieldType;
                case MemberTypes.Property:
                    return ((PropertyInfo)member).PropertyType;
                case MemberTypes.Method:
                    return ((MethodInfo)member).ReturnType;
                default:
                    throw new InvalidOperationException($"LoadFromCollection: Unsupported MemberType on member {member.Name}. Only Field, Property and Method allowed.");
            }
        }
    }
}
