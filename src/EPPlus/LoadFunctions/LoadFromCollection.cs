using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeOpenXml.LoadFunctions.Params;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromCollection<T> : LoadFunctionBase
    {
        public LoadFromCollection(ExcelRangeBase range, IEnumerable<T> items, LoadFromCollectionParams parameters) : base(range, parameters)
        {
            _items = items;
            _members = parameters.Members;
            _bindingFlags = parameters.BindingFlags;
            _headerParsingType = parameters.HeaderParsingType;
            var type = typeof(T);
            if (_members == null)
            {
                _members = type.GetProperties(_bindingFlags);
            }
            else
            {
                if (_members.Length == 0)   //Fixes issue 15555
                {
                    throw (new ArgumentException("Parameter Members must have at least one property. Length is zero"));
                }
                foreach (var t in _members)
                {
                    if (t.DeclaringType != null && t.DeclaringType != type)
                    {
                        _isSameType = false;
                    }
                    //Fixing inverted check for IsSubclassOf / Pullrequest from tomdam
                    if (t.DeclaringType != null && t.DeclaringType != type && !TypeCompat.IsSubclassOf(type, t.DeclaringType) && !TypeCompat.IsSubclassOf(t.DeclaringType, type))
                    {
                        throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
                    }
                }
            }
        }

        private readonly BindingFlags _bindingFlags;
        private readonly MemberInfo[] _members;
        private readonly HeaderParsingTypes _headerParsingType;
        private readonly IEnumerable<T> _items;
        private readonly bool _isSameType;

        protected override int GetNumberOfColumns()
        {
            return _members.Length == 0 ? 1 : _members.Length;
        }

        protected override int GetNumberOfRows()
        {
            if (_items == null) return 0;
            return _items.Count();
        }

        protected override void LoadInternal(object[,] values)
        {

            int col = 0, row = 0;
            if (_members.Length > 0 && PrintHeaders)
            {
                foreach (var t in _members)
                {
                    var descriptionAttribute = t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() as DescriptionAttribute;
                    var header = string.Empty;
                    if (descriptionAttribute != null)
                    {
                        header = descriptionAttribute.Description;
                    }
                    else
                    {
                        var displayNameAttribute =
                            t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as
                            DisplayNameAttribute;
                        if (displayNameAttribute != null)
                        {
                            header = displayNameAttribute.DisplayName;
                        }
                        else
                        {
                            header = ParseHeader(t.Name);
                        }
                    }
                    //_worksheet.SetValueInner(row, col++, header);
                    values[row, col++] = header;
                }
                row++;
            }

            if (!_items.Any() && (_members.Length == 0 || PrintHeaders == false))
            {
                return;
            }

            var nMembers = GetNumberOfColumns();
            foreach (var item in _items)
            {
                if (item == null)
                {
                    col = GetNumberOfColumns();
                }
                else
                {
                    col = 0;
                    if (item is string || item is decimal || item is DateTime || TypeCompat.IsPrimitive(item))
                    {
                        values[row, col++] = item;
                    }
                    else
                    {
                        foreach (var t in _members)
                        {
                            if (_isSameType == false && item.GetType().GetMember(t.Name, _bindingFlags).Length == 0)
                            {
                                col++;
                                continue; //Check if the property exists if and inherited class is used
                            }
                            else if (t is PropertyInfo)
                            {
                                values[row, col++] = ((PropertyInfo)t).GetValue(item, null);
                            }
                            else if (t is FieldInfo)
                            {
                                values[row, col++] = ((FieldInfo)t).GetValue(item);
                            }
                            else if (t is MethodInfo)
                            {
                                values[row, col++] = ((MethodInfo)t).Invoke(item, null);
                            }
                        }
                    }
                }
                row++;
            }
        }

        private string ParseHeader(string header)
        {
            switch(_headerParsingType)
            {
                case HeaderParsingTypes.Preserve:
                    return header;
                case HeaderParsingTypes.UnderscoreToSpace:
                    return header.Replace("_", " ");
                case HeaderParsingTypes.CamelCaseToSpace:
                    return Regex.Replace(header, "([A-Z])", " $1", RegexOptions.Compiled).Trim();
                case HeaderParsingTypes.UnderscoreAndCamelCaseToSpace:
                    header = Regex.Replace(header, "([A-Z])", " $1", RegexOptions.Compiled).Trim();
                    return header.Replace("_ ", "_").Replace("_", " ");
                default:
                    return header;
            }
        }
    }
}

