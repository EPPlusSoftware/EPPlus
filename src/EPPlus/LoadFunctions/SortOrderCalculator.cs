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

namespace OfficeOpenXml.LoadFunctions
{
    internal class SortOrderCalculator
    {
        public SortOrderCalculator(NestedColumnsTypeScanner scanner, MemberInfo[] membersFilter)
        {
            _membersFilter= membersFilter != null ? membersFilter.ToList() : null;
            var allTypes = scanner.GetTypes();
            foreach (var type in allTypes)
            {
                if(type.HasAttributeOfType<EPPlusTableColumnSortOrderAttribute>())
                {
                    var attr = type.GetFirstAttributeOfType<EPPlusTableColumnSortOrderAttribute>();
                    if(!_sortOrderAttributes.ContainsKey(type))
                    {
                        _sortOrderAttributes[type] = attr.Properties.ToList();
                    }
                }
            }
        }

        private readonly Dictionary<Type, List<string>> _sortOrderAttributes = new Dictionary<Type, List<string>>();
        private readonly List<MemberInfo> _membersFilter;

        public void CalculateSortOrder(
            ref List<int> sortOrderList, 
            int memberIndex, 
            int nestedLevel, 
            MemberPath memberPath,
            MemberInfo member)
        {
            if(sortOrderList == null) sortOrderList = new List<int>();
            var sortOrder = memberIndex;
            var declaringType = member.DeclaringType;
            if(_membersFilter != null && _membersFilter.Contains(member))
            {
                sortOrderList.Add(_membersFilter.IndexOf(member));
            }
            else if (member.HasAttributeOfType<EpplusNestedTableColumnAttribute>())
            {
                var attr = member.GetFirstAttributeOfType<EpplusNestedTableColumnAttribute>();
                if (attr.Order > 0)
                {
                    sortOrder = attr.Order;
                }
            }
            else if (_sortOrderAttributes.ContainsKey(declaringType))
            {
                sortOrder = _sortOrderAttributes[declaringType].IndexOf(member.Name);
            }
            else if(member.HasAttributeOfType<EpplusTableColumnAttribute>())
            {
                var attr = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
                if(attr.Order > 0)
                {
                    sortOrder = attr.Order;
                }
            }
            sortOrderList.Add(sortOrder);
        }
    }
}
