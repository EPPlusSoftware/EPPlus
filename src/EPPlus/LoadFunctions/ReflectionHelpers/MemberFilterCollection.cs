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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    internal class MemberFilterCollection : IEnumerable<MemberInfo>
    {
        public MemberFilterCollection(MemberInfo[] members = null)
        {
            if(members != null)
            {
                _members = new List<MemberInfo>(members);
            }
            else
            {
                _members = new List<MemberInfo>();
            }
        }

        private readonly List<MemberInfo> _members;

        public bool IsEmpty => _members.Any() == false;

        public bool Exists(MemberInfo member)
        {
            if (IsEmpty) return false;
            foreach (var filter in _members)
            {
                if (filter.DeclaringType == member.DeclaringType && filter.Name == member.Name) return true;
                if (filter.DeclaringType == member.GetMemberType()) return true;
            }
            return false;
        }

        public List<MemberInfo> ToList()
        {
            return _members;
        }

        public int GetNumberOfChildrenByParent(MemberPath parentPath)
        {
            if (parentPath == null) return 0;
            var parentType = parentPath.Last().Member.GetMemberType();
            return _members.Count(x => x.DeclaringType == parentType);
        }
        
        public bool IsOnlyChild(MemberPath parentPath, MemberInfo child)
        {
            if(parentPath == null) return false;
            var parentType = parentPath.Last().Member.GetMemberType();
            var children = _members.Where(x => x.DeclaringType == parentType);
            return children.Count() == 1 && children.First().Name == child.Name;
        }

        public IEnumerator<MemberInfo> GetEnumerator()
        {
            return ((IEnumerable<MemberInfo>)_members).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_members).GetEnumerator();
        }
    }
}
