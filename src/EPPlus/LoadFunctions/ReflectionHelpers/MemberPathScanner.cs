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
using OfficeOpenXml.LoadFunctions.Params;

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
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
            _filterMembers = parameters.Members;
            _params = parameters;
            if (_filterMembers != null && _filterMembers.Length > 0)
            {
                var usedTypesScanner = new UsedTypesScanner(outerType);
                usedTypesScanner.ValidateMembers(_filterMembers);
            }
            ReadTypes(outerType);
        }

        private readonly MemberInfo[] _filterMembers;
        private readonly LoadFromCollectionParams _params;
        private readonly List<MemberPath> _paths = new List<MemberPath>();

        private bool ShouldAddPath(MemberPath parentPath, MemberInfo member)
        {
            if (member.HasAttributeOfType<EpplusIgnore>()) return false;
            if (_filterMembers == null || _filterMembers.Length == 0) return true;
            if (member.ExistsInFilter(_filterMembers) == false) return false;
            if (parentPath?.Last().IsNestedProperty ?? false)
            {
                var parentMember = parentPath.Last().Member;

                // If the only filtered member is a complex type
                // with the EpplusNestedTableColumn we should
                // use all its members.
                if (_filterMembers.Any(x => x.Name == parentMember.Name && x.DeclaringType == parentMember.DeclaringType))
                {
                    // if there are members specified explicitly in the filter
                    // we should only use those.
                    if (_filterMembers.Count(x => x.DeclaringType == parentMember.DeclaringType) > 1)
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
            var parentIsNested = path != null && path.Depth > 0 && path.Last().IsNestedProperty;
            var index = 0;
            var members = type.GetLoadFromCollectionMembers(_params.BindingFlags, _filterMembers);
            foreach (var member in members)
            {
                var mType = member.GetMemberType();
                var shouldAddPath = ShouldAddPath(path, member);
                if (parentIsNested == false && shouldAddPath == false)
                {
                    continue;
                }
                else if(shouldAddPath == false)
                {

                }
                var sortOrder = member.GetSortOrder(_filterMembers, index, out bool useForAllPathItems);
                var propPath = MemberPath.CreateNewOrAppend(path, member, sortOrder, useForAllPathItems);
                var lastItem = propPath.Last();
                if (member.HasAttributeOfType(out EpplusNestedTableColumnAttribute entAttr))
                {
                    var memberType = member.GetMemberType().GetTypeOrUnderlyingType();
                    if (memberType == typeof(string) || !memberType.IsClass && memberType.IsInterface)
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
                    var dictPaths = DictionaryColumnPathFactory.Create(member, propPath, _params);
                    _paths.AddRange(dictPaths);
                }
                if (member.HasAttributeOfType(out EpplusTableColumnAttribute etcAttr))
                {
                    lastItem.SetProperties(etcAttr);
                }
                if (lastItem.IsNestedProperty == false && lastItem.IsDictionaryParent == false)
                {
                    _paths.Add(propPath);
                }
                index++;
            }
        }

        public List<MemberPath> GetPaths()
        {
            return _paths;
        }
    }
}
