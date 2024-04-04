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
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.LoadFunctions.ReflectionHelpers;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal static class DictionaryColumnPathFactory
    {
        public static IEnumerable<MemberPath> Create(MemberInfo member, MemberPath path, LoadFromCollectionParams parameters)
        {
            var result = new List<MemberPath>();
            var attr = member.GetFirstAttributeOfType<EPPlusDictionaryColumnAttribute>();
            var memberType = member.GetMemberType();
            var memberPath = path.GetPath();
            if (memberType != typeof(Dictionary<string, object>))
            {
                throw new InvalidOperationException($"Property {memberPath} is decorated with the EPPlusDictionaryColumnsAttribute. Its type must be Dictionary<string, object>");
            }
            var columnHeaders = Enumerable.Empty<string>();
            if (!string.IsNullOrEmpty(attr.KeyId))
            {
                columnHeaders = parameters.GetDictionaryKeys(attr.KeyId);
            }
            else if (attr.ColumnHeaders != null && attr.ColumnHeaders.Length > 0)
            {
                columnHeaders = attr.ColumnHeaders;
            }
            else
            {
                columnHeaders = parameters.GetDefaultDictionaryKeys();
            }
            var index = 0;
            foreach (var key in columnHeaders)
            {
                var propPath = path.Clone();
                var itemMember = new DictionaryItemMemberInfo(key, member);
                var item = new MemberPathItem(itemMember, key, index++);
                propPath.Append(item);
                result.Add(propPath);
            }
            return result;
        }
    }
}
