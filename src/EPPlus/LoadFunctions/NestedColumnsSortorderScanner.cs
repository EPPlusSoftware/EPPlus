using OfficeOpenXml.Attributes;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class NestedColumnsSortorderScanner
    {
        public NestedColumnsSortorderScanner(Type outerType, BindingFlags bindingFlags)
        {
            _sortOrder = new List<string>();
            _bindingFlags = bindingFlags;
            ReadTypes(outerType);
        }

        private readonly List<string> _sortOrder = new List<string>();
        private readonly BindingFlags _bindingFlags;

        private Type GetMemberType(MemberInfo member, string path)
        {
            if(member.MemberType == MemberTypes.Property)
            {
                return ((PropertyInfo)member).PropertyType;
            }
            else if(member.MemberType == MemberTypes.Field)
            {
                return ((FieldInfo)member).FieldType;
            }
            else if(member.MemberType == MemberTypes.Method)
            {
                return ((MethodInfo)member).ReturnType;
            }
            else
            {
                throw new InvalidCastException($"Member {path} must be either Property, Field or Method but is {member.MemberType} which is not supported.");
            }
        }

        private string GetMemberName(MemberInfo member)
        {
            if (member.MemberType == MemberTypes.Property)
            {
                return ((PropertyInfo)member).Name;
            }
            else if (member.MemberType == MemberTypes.Field)
            {
                return ((FieldInfo)member).Name;
            }
            else if (member.MemberType == MemberTypes.Method)
            {
                return ((MethodInfo)member).Name;
            }
            else
            {
                throw new InvalidCastException($"Member {member.Name} must be either Property, Field or Method but is {member.MemberType} which is not supported.");
            }
        }

        private void ReadTypes(Type type, string path = null)
        {
            if(type.HasPropertyOfType<EPPlusTableColumnSortOrderAttribute>())
            {
                var sortOrderAttribute = type.GetFirstAttributeOfType<EPPlusTableColumnSortOrderAttribute>();
                if(_sortOrder.Count == 0)
                {
                    _sortOrder.AddRange(sortOrderAttribute.Properties);
                }
                else
                {
                    var pathIndex = _sortOrder.IndexOf(path);
                    var offset = 1;
                    foreach(var prop in sortOrderAttribute.Properties)
                    {
                        var fullPath = $"{path}.{prop}";
                        _sortOrder.Insert(pathIndex + offset++, fullPath);
                    }
                    _sortOrder.Remove(path);
                }
                foreach (var member in type.GetProperties(_bindingFlags))
                {
                    if (member.MemberType != MemberTypes.Property) continue;
                    var memberName = GetMemberName(member);
                    var memberPath = string.IsNullOrEmpty(path) ? member.Name : $"{path}.{memberName}";
                    var isNested = member.HasPropertyOfType<EpplusNestedTableColumnAttribute>();
                    if(isNested)
                    {
                        var memberType = GetMemberType(member, memberPath);
                        ReadTypes(memberType, memberPath);
                    }
                }

            }
        }

        public List<string> GetSortOrder()
        {
            return _sortOrder;
        }
    }
}
