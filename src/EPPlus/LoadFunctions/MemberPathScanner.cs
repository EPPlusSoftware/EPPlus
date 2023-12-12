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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions.ReflectionHelpers;
using OfficeOpenXml.LoadFunctions.Params;

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// Scans a type for properties decorated with the <see cref="EpplusNestedTableColumnAttribute"/>
    /// and returns a list with all types reflected by these properties including the outer type.
    /// </summary>
    internal class MemberPathScanner
    {
        public MemberPathScanner(
            Type outerType, 
            LoadFromCollectionParams parameters)
        {
            _bindingFlags = parameters.BindingFlags;
            _filterMembers= parameters.Members;
            _params = parameters;
            if(_filterMembers != null && _filterMembers.Length > 0)
            {
                var usedTypesScanner = new UsedTypesScanner(outerType);
                usedTypesScanner.ValidateMembers(_filterMembers);
            }
            ReadTypes(outerType);
        }

        private readonly BindingFlags _bindingFlags;
        private readonly MemberInfo[] _filterMembers;
        private readonly LoadFromCollectionParams _params;
        private readonly List<MemberPath> _paths = new List<MemberPath>();

        private bool NestedClassExistsInFilter(MemberInfo nestedClass)
        {
            if(_filterMembers == null || _filterMembers.Length == 0) return false;
            foreach(var filter in  _filterMembers)
            {
                if(filter.DeclaringType == nestedClass.DeclaringType && filter.Name == nestedClass.Name) return true;
                if (filter.DeclaringType == nestedClass.GetMemberType()) return true;
            }
            return false;
        }

        private bool ShouldAddPath(MemberPath parentPath, MemberInfo member)
        {
            if (member.HasAttributeOfType<EpplusIgnore>()) return false;
            if (_filterMembers == null || _filterMembers.Length == 0) return true;
            if (!NestedClassExistsInFilter(member)) return false;
            if(parentPath?.Last().IsNestedProperty ?? false)
            {
                var parentMember = parentPath.Last().Member;

                // If the only filtered member is a complex type
                // with the EpplusNestedTableColumn we should
                // use all its members.
                if (_filterMembers.Any(x => x.Name == parentMember.Name && x.DeclaringType == parentMember.DeclaringType))
                {
                    // if there are members specified explicitly in the filter
                    // we should only use those.
                    if(_filterMembers.Count(x => x.DeclaringType == parentMember.DeclaringType) > 1)
                    {
                        return _filterMembers.Any(x => x.Name == member.Name && x.DeclaringType == parentMember.DeclaringType);
                    }
                    return true;
                }
            }
            // Always ignore complex type members not decorated with the
            // EpplusNestedTableColumn attribute.
            if (member.HasAttributeOfType<EpplusNestedTableColumnAttribute>()) return true;
            return member.GetMemberType().IsComplexType() == false;
        }

        private void ReadTypes(Type type, MemberPath path = null)
        {
            var members = type.GetProperties(_bindingFlags).Where(x => x.ShouldBeIncluded());
            var parentIsNested = path != null && path.Depth > 0 && path.Last().IsNestedProperty;
            var index = 0;
            foreach(var member in members)
            {
                var mType = member.GetMemberType();
                if(parentIsNested == false && ShouldAddPath(path, member) == false)
                {
                    continue;
                }
                var sortOrder = index;
                var calculatedSortOrder = member.GetSortOrder(_filterMembers, out bool useForAllPathItems);
                if(calculatedSortOrder.HasValue)
                {
                    // some attributes has int.MaxValue as default value
                    // so means that order hasn't been set.
                    sortOrder = calculatedSortOrder.Value == int.MaxValue ?
                        ExcelPackage.MaxColumns + index
                        :
                        calculatedSortOrder.Value;
                }
                var propPath = path?.Clone();
                if (propPath == null)
                {
                    propPath = new MemberPath(member, sortOrder, useForAllPathItems);
                }
                else
                {
                    propPath.Append(member, sortOrder, useForAllPathItems);
                }

                var lastItem = propPath.Last();
                if (member.HasAttributeOfType(out EpplusNestedTableColumnAttribute entAttr))
                {
                    var memberType = member.GetMemberType();
                    if (memberType == typeof(string) || (!memberType.IsClass && memberType.IsInterface))
                    {
                        throw new InvalidOperationException($"EpplusNestedTableColumn attribute can only be used with complex types (member: {propPath.GetPath()})");
                    }
                    lastItem.SetProperties(entAttr);
                    ReadTypes(mType, propPath);
                }
                if (member.HasAttributeOfType(out EPPlusDictionaryColumnAttribute edcAttr))
                {
                    lastItem.IsDictionaryParent = true;
                    lastItem.DictionaryKey = edcAttr.KeyId;
                    var dictPaths = HandleDictionaryColumnsAttribute(member, propPath);
                    _paths.AddRange(dictPaths);
                }
                if(member.HasAttributeOfType(out EpplusTableColumnAttribute etcAttr))
                {
                    lastItem.SetProperties(etcAttr);
                }
                if(lastItem.IsNestedProperty == false && lastItem.IsDictionaryParent == false) 
                {
                    _paths.Add(propPath);
                }
                index++;
            }
        }

        private IEnumerable<MemberPath> HandleDictionaryColumnsAttribute(MemberInfo member, MemberPath path)
        {
            var result = new List<MemberPath>();
            var attr = member.GetFirstAttributeOfType<EPPlusDictionaryColumnAttribute>();
            var memberType = member.GetMemberType();
            var memberPath = path.GetPath();
            if(memberType != typeof(Dictionary<string, object>))
            {
                throw new InvalidOperationException($"Property {memberPath} is decorated with the EPPlusDictionaryColumnsAttribute. Its type must be Dictionary<string, object>");
            }
            var columnHeaders = Enumerable.Empty<string>();
            if (!string.IsNullOrEmpty(attr.KeyId))
            {
                columnHeaders = _params.GetDictionaryKeys(attr.KeyId);
            }
            else if (attr.ColumnHeaders != null && attr.ColumnHeaders.Length > 0)
            {
                columnHeaders = attr.ColumnHeaders;
            }
            else
            {
                columnHeaders = _params.GetDefaultDictionaryKeys();
            }
            var index = 0;
            foreach (var key in columnHeaders)
            {
                //var sortOrderListCol = CopyList(sortOrderList);
                // _sortOrderCalculator.CalculateSortOrder(ref sortOrderListCol, index++, nestedLevel, memberPath, member);
                //result.Add(new ColumnInfo
                //{
                //    Index = index,
                //    MemberInfo = member,
                //    IsDictionaryProperty = true,
                //    DictinaryKey = key,
                //    //Path = $"{memberPath}.{key}",
                //    Header = key,
                //    //SortOrderLevels = sortOrderListCol
                //});
                //path.
                var propPath = path.Clone();
                var itemMember = new DictionaryItemMemberInfo(key);
                var item = new MemberPathItem(itemMember, key, index++);
                propPath.Append(item);
                result.Add(propPath);
            }
            return result;
        }

        /// <summary>
        /// Returns all the scanned types, including the outer type
        /// </summary>
        /// <returns></returns>
        //public HashSet<Type> GetTypes()
        //{
        //    return _types;
        //}

        public List<MemberPath> GetPaths()
        {
            return _paths;
        }

        /// <summary>
        /// Returns true if the <paramref name="type"/> exists among the scanned types.
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        //public bool Exists(Type type)
        //{
        //    return _types.Contains(type);
        //}
    }
}
