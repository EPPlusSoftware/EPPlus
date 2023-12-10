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
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    [DebuggerDisplay("Path: {GetPath()}, IsNested: {Last().IsNestedProperty}")]
    internal class MemberPath
    {
        public MemberPath()
        {

        }
        public MemberPath(MemberInfo member)
        {
            _members.Add(new MemberPathItem(member));
        }

        private readonly List<MemberPathItem> _members = new List<MemberPathItem>();

        internal void Append(MemberInfo member)
        {
            _members.Add(new MemberPathItem(member));
        }

        internal string GetPath()
        {
            var members = _members.Select(m => m.Member.Name).ToList();
            return string.Join(".", members.ToArray());
        }

        public int Depth
        {
            get => _members.Count;
        }

        public MemberPathItem Get(int index)
        {
            if (index >= _members.Count) throw new IndexOutOfRangeException();
            return _members[index];
        }

        public bool IsParentTo(MemberPath other)
        {
            if (other.Depth != Depth - 1) return false;
            for(var x = 0; x < other.Depth;x++)
            {
                var thisVal = _members[x].Member.Name;
                var otherVal = other.Get(x).Member.Name;
                if(string.Compare(thisVal, otherVal, true) != 0) return false;
            }
            return true;
        }

        public MemberPath Clone()
        {
            var c = new MemberPath();
            for(var x = 0; x < _members.Count;x++)
            {
                c.Append(Get(x).Member);
            }
            return c;
        }

        public bool IsChildTo(MemberPath other)
        {
            if (other.Depth != Depth + 1) return false;
            for (var x = 0; x < _members.Count; x++)
            {
                var thisVal = _members[x].Member.Name;
                var otherVal = other.Get(x).Member.Name;
                if (string.Compare(thisVal, otherVal, true) != 0) return false;
            }
            return true;
        }

        public MemberPathItem Last()
        {
            return _members.Last();
        }
    }
}
