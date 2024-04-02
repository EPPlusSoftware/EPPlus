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
            if(parameters.Members != null)
            {
                _filterMembers = new MemberFilterCollection(parameters.Members);
                var usedTypesScanner = new UsedTypesScanner(outerType);
                usedTypesScanner.ValidateMembers(_filterMembers);
            }
            _params = parameters;
            Scan(outerType);
        }

        private readonly MemberFilterCollection _filterMembers = new MemberFilterCollection();
        private readonly LoadFromCollectionParams _params;
        private readonly List<MemberPath> _paths = new List<MemberPath>();

        private bool ShouldAddPath(MemberPath parentPath, MemberInfo member)
        {
            if (member.HasAttributeOfType<EpplusIgnore>()) return false;
            if (_filterMembers.IsEmpty) return true;
            if (parentPath?.Last().IsNestedProperty ?? false)
            {
                // if the parent is a nested property with no child properties
                // all child properties should be included.
                var nChildren = _filterMembers.GetNumberOfChildrenByParent(parentPath);
                if (nChildren == 0) return true;
            }
            if (_filterMembers.Exists(member) == false) return false;
            // Always ignore complex type members not decorated with the
            // EpplusNestedTableColumn attribute.
            if (member.HasAttributeOfType<EpplusNestedTableColumnAttribute>()) return true;
            return member.GetMemberType().IsComplexType() == false;
        }

        private void Scan(Type type, MemberPath path = null)
        {
            var parentIsNested = path != null && path.Depth > 0 && path.Last().IsNestedProperty;
            var index = 0;
            var members = type.GetLoadFromCollectionMembers(_params.BindingFlags, _filterMembers);
            foreach (var member in members)
            {
                var mType = member.GetMemberType();
                var shouldAddPath = ShouldAddPath(path, member);
                if (shouldAddPath == false)
                {
                    continue;
                }
                var sortOrder = member.GetSortOrder(_filterMembers.ToList(), index, out bool useForAllPathItems);
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
                    Scan(mType, propPath);
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
