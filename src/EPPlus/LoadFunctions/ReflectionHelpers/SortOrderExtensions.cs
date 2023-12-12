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
using OfficeOpenXml.Utils;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
#if !NET35
using System.ComponentModel.DataAnnotations;
#endif

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    internal static class SortOrderExtensions
    {
        public static int? GetSortOrder(this MemberInfo member, MemberInfo[] filterMembers, out bool useForAllPathItems)
        {
            useForAllPathItems = false;
            if(filterMembers != null && filterMembers.Length > 0)
            {
                for(int i = 0; i < filterMembers.Length; i++)
                {
                    var m = filterMembers[i];
                    if(m.MemberType == member.MemberType 
                        && m.DeclaringType == member.DeclaringType
                        && m.Name == member.Name
                        )
                    {
                        useForAllPathItems = true;
                        return i;
                    }
                }
            }
            if(member.DeclaringType.HasAttributeOfType<EPPlusTableColumnSortOrderAttribute>())
            {
                var attr = member.DeclaringType.GetFirstAttributeOfType<EPPlusTableColumnSortOrderAttribute>();
                return attr.Properties.ToList().IndexOf(member.Name);
            }
            if (member.HasAttributeOfType(out EpplusNestedTableColumnAttribute entcAttr))
            {
                return entcAttr.Order;
            }
            if(member.HasAttributeOfType(out EpplusTableColumnAttribute etcAttr))
            {
                return etcAttr.Order;
            }
            if (member.HasAttributeOfType(out EPPlusDictionaryColumnAttribute edcAttr))
            {
                return edcAttr.Order;
            }
#if !NET35
            if(member.HasAttributeOfType(out DisplayAttribute displayAttr))
            {
                return displayAttr.Order;
            }
#endif
            return default;
        }
    }
}
