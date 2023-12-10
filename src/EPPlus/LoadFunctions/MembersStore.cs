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
    internal class MembersStore<T>
    {
        public MembersStore(
            MemberInfo[] filterMembers,
            BindingFlags bindingFlags
            )
        {
            _bindingFlags = bindingFlags;
            _typeScanner = new NestedColumnsTypeScanner(typeof(T), filterMembers, bindingFlags);
            _includedTypes = new HashSet<Type>(_typeScanner.GetTypes().Distinct());
            if (filterMembers != null)
            {
                _filterMembers = filterMembers.ToList();
            }
            _members = new Dictionary<Type, HashSet<string>>();
            if (filterMembers != null && filterMembers.Length > 0)
            {
                foreach (var member in filterMembers)
                {
                    AddMember(member);
                }
            }
        }

        private readonly Dictionary<Type, HashSet<string>> _members;
        private readonly HashSet<Type> _includedTypes;
        private List<MemberInfo> _filterMembers;
        private readonly BindingFlags _bindingFlags;
        private readonly NestedColumnsTypeScanner _typeScanner;

        private bool SubMemberIsSpecified(MemberInfo member)
        {
            if(_filterMembers == null || _filterMembers.Count == 0)
            {
                return false;
            }
            else if(_members.ContainsKey(member.DeclaringType))
            {
                return _filterMembers.Contains(member);
            }
            return false;
        }

        internal NestedColumnsTypeScanner GetTypeScanner()
        {
            return _typeScanner;
        }

        internal void ValidateType(MemberInfo member)
        {
            MemberHelper.ValidateType(member, _includedTypes);
        }

        internal void AddMember(MemberInfo member)
        {
            if (!_members.ContainsKey(member.DeclaringType))
            {
                _members.Add(member.DeclaringType, new HashSet<string>());
            }

            _members[member.DeclaringType].Add(member.Name);
        }

        internal bool ShouldIgnoreMember(MemberInfo member, MemberPath path)
        {
            if (member == null) return true;
            var isNested = path.Depth > 1;
            if (member.HasAttributeOfType<EpplusIgnore>()) return true;
            if (_members.Count == 0) return false;
            if (isNested && (_members == null || !_members.ContainsKey(member.DeclaringType)))
            {
                return false;
            }
            return !(_members.ContainsKey(member.DeclaringType) && _members[member.DeclaringType].Contains(member.Name));
        }

        internal IEnumerable<MemberInfo> GetMembers(Type type, int nestedLevel)
        {
            return nestedLevel == 0 && _filterMembers != null ? _filterMembers.ToArray() : type.GetProperties(_bindingFlags);
        }

    }
}
