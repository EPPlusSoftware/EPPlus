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

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// Scans a type for properties decorated with the <see cref="EpplusNestedTableColumnAttribute"/>
    /// and returns a list with all types reflected by these properties including the outer type.
    /// </summary>
    internal class NestedColumnsTypeScanner
    {
        public NestedColumnsTypeScanner(Type outerType, MemberInfo[] filterMembers, BindingFlags bindingFlags)
        {
            _bindingFlags = bindingFlags;
            _filterMembers= filterMembers;
            _types.Add(outerType);
            ReadTypes(outerType);
        }

        private readonly HashSet<Type> _types = new HashSet<Type>();
        private readonly BindingFlags _bindingFlags;
        private readonly MemberInfo[] _filterMembers;
        private readonly List<MemberPath> _paths = new List<MemberPath>();

        private void ReadTypes(Type type, bool isNested = false, MemberPath path = null)
        {
            var properties = type.GetProperties(_bindingFlags);
            foreach(var property in properties)
            {
                if (property.HasAttributeOfType<EpplusIgnore>()) continue;
                var propPath = path?.Clone();
                if(propPath == null)
                {
                    propPath = new MemberPath(property);
                }
                else
                {
                    propPath.Append(property);
                }
                if (property.HasAttributeOfType<EpplusNestedTableColumnAttribute>())
                {
                    propPath.Last().IsNestedProperty = true;
                    if(!_types.Contains(property.PropertyType))
                    {
                        _types.Add(property.PropertyType);
                        ReadTypes(property.PropertyType, true, propPath);
                    }
                }
                if (
                    _filterMembers == null ||
                    _filterMembers.Length == 0 ||
                    _filterMembers.Any(x => x.Name == property.Name && x.DeclaringType == property.DeclaringType) ||
                    isNested
                    )
                {
                    _paths.Add(propPath);
                }
                    
            }
        }

        /// <summary>
        /// Returns all the scanned types, including the outer type
        /// </summary>
        /// <returns></returns>
        public HashSet<Type> GetTypes()
        {
            return _types;
        }

        /// <summary>
        /// Returns true if the <paramref name="type"/> exists among the scanned types.
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public bool Exists(Type type)
        {
            return _types.Contains(type);
        }
    }
}
