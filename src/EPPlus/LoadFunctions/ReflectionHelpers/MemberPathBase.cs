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

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    internal abstract class MemberPathBase
    {
        public abstract string GetHeader();

        protected readonly List<MemberPathItem> _members = new List<MemberPathItem>();

        public MemberPathItem Last()
        {
            return _members.Last();
        }

        public object GetLastMemberValue(object item, BindingFlags bindingFlags)
        {
            if (IsFormulaColumn) return null;
            object v = item;
            for (var i = 0; i < Depth && v != null; i++)
            {
                var pathItem = _members[i];
                v = pathItem.Member.GetValue(v, bindingFlags);
            }
            return v;
        }

        public virtual bool IsFormulaColumn { get; } = false;

        public int Depth
        {
            get => _members.Count;
        }

        public MemberPathItem Get(int index)
        {
            if (index >= _members.Count) throw new IndexOutOfRangeException();
            return _members[index];
        }

        internal virtual string GetPath()
        {
            var members = _members.Select(m => m.Member.Name).ToList();
            return string.Join(".", members.ToArray());
        }

        public MemberPath Clone()
        {
            var c = new MemberPath();
            for (var x = 0; x < _members.Count; x++)
            {
                var item = Get(x).Clone();
                c.Append(item);
            }
            return c;
        }

        internal string GetSortOrderString()
        {
            var orders = _members.Select(m => m.SortOrder.ToString()).ToList();
            return string.Join(",", orders.ToArray());
        }
    }
}
