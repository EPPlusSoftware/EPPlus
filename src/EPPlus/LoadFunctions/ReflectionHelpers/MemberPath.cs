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
using System.ComponentModel;
#if !NET35
using System.ComponentModel.DataAnnotations;
#endif
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeOpenXml.LoadFunctions.Params;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    [DebuggerDisplay("Path: {GetPath()}, IsNested: {Last().IsNestedProperty}, Order={GetSortOrderString()}")]
    internal class MemberPath : MemberPathBase
    {
        public MemberPath()
        {

        }
        public MemberPath(MemberInfo member, int sortOrder, bool useForAllPathItems = false)
        {
            var newItem = new MemberPathItem(member, sortOrder);
            if (_members.Any())
            {
                newItem.Parent = _members.Last();
            }
            _members.Add(newItem);
            if (useForAllPathItems)
            {
                _members.ForEach(x => x.SortOrder = sortOrder);
            }
        }

        internal void Append(MemberPathItem item, bool useSortOrderForAllPathItems = false)
        {
            if (_members.Any())
            {
                item.Parent = _members.Last();
            }
            _members.Add(item);
            if (useSortOrderForAllPathItems)
            {
                _members.ForEach(x => x.SortOrder = item.SortOrder);
            }
        }

        internal void Append(MemberInfo member, int sortOrder, bool useForAllPathItems = false)
        {
            var newItem = new MemberPathItem(member, sortOrder);
            if (_members.Any())
            {
                newItem.Parent = _members.Last();
            }
            _members.Add(newItem);
            if (useForAllPathItems)
            {
                _members.ForEach(x => x.SortOrder = sortOrder);
            }
        }

        public override string GetHeader()
        {
            string prefix = string.Empty;
            string header = string.Empty;
            List<string> prefixes = new List<string>();
            var last = _members.Last();
            var tmp = last;
            while (tmp.Parent != null)
            {
                tmp = tmp.Parent;
                if (!string.IsNullOrEmpty(tmp.HeaderPrefix))
                {
                    prefixes.Insert(0, tmp.HeaderPrefix);
                }
            }
            if (prefixes.Count > 0)
            {
                prefix = string.Join(" ", prefixes.ToArray());
            }
            if (last.IsDictionaryColumn)
            {
                header = last.DictionaryKey;
            }
            else if (last.Member.HasAttributeOfType(out EpplusTableColumnAttribute etcAttr) && !string.IsNullOrEmpty(etcAttr.Header))
            {
                header = etcAttr.Header;
            }
            else if (last.Member.HasAttributeOfType(out DescriptionAttribute descAttr))
            {
                header = descAttr.Description;
            }
            else if (last.Member.HasAttributeOfType(out DisplayNameAttribute displayNameAttr))
            {
                header = displayNameAttr.DisplayName;
            }
#if !NET35
            else if (last.Member.HasAttributeOfType(out DisplayAttribute displayAttr))
            {
                header = displayAttr.Name;
            }
#endif
            if (string.IsNullOrEmpty(header))
            {
                header = last.Member.Name;
            }
            if (!string.IsNullOrEmpty(prefix))
            {
                return $"{prefix} {header}";
            }
            return header;
        }

        public static MemberPath CreateNewOrAppend(MemberPath path, MemberInfo member, int sortOrder, bool useSortOrderForAllPathItems)
        {
            var resultPath = path?.Clone();
            if (resultPath == null)
            {
                resultPath = new MemberPath(member, sortOrder, useSortOrderForAllPathItems);
            }
            else
            {
                resultPath.Append(member, sortOrder, useSortOrderForAllPathItems);
            }
            return resultPath;
        }
    }
}
