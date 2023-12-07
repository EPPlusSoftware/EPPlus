/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/286/2021         EPPlus Software AB       EPPlus 5.7.5
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromCollectionColumns<T>
    {
        public LoadFromCollectionColumns(LoadFromCollectionParams parameters):
            this(parameters, Enumerable.Empty<string>().ToList())
        { }

        public LoadFromCollectionColumns(LoadFromCollectionParams parameters, List<string> sortOrderColumns)
        {
            _params = parameters;
            _bindingFlags = parameters.BindingFlags;
            _sortOrderColumns = sortOrderColumns;
            if(parameters.Members != null)
            {
                _filterMembers = parameters.Members.ToList();
            }
            var typeScanner = new NestedColumnsTypeScanner(typeof(T), _bindingFlags);
            _sortOrderCalculator = new SortOrderCalculator(typeScanner);
            _includedTypes = new HashSet<Type>(typeScanner.GetTypes().Distinct());
            var members = parameters.Members;
            _members = new Dictionary<Type, HashSet<string>>();
            if (members != null && members.Length > 0)
            {
                foreach (var member in members)
                {
                    AddMember(member);
                }
            }
        }


        private readonly LoadFromCollectionParams _params;
        private readonly BindingFlags _bindingFlags;
        private readonly List<string> _sortOrderColumns;
        private readonly SortOrderCalculator _sortOrderCalculator;
        private readonly Dictionary<Type, HashSet<string>> _members;
        private List<MemberInfo> _filterMembers;
        private readonly HashSet<Type> _includedTypes;
        private const int SortOrderOffset = ExcelPackage.MaxColumns;
        private int _currentColumn = 0;

        internal List<ColumnInfo> Setup()
        {
            var result = new List<ColumnInfo>();
            var t = typeof(T);
            var ut = Nullable.GetUnderlyingType(t);
            if (ut != null)
            {
                t = ut;
            }

            bool sort=SetupInternal(t, result, null);
            if (sort)
            {
                ReindexAndSortColumns(result);
            }
            return result;
        }

        private void AddMember(MemberInfo member)
        {
            if (!_members.ContainsKey(member.DeclaringType))
            {
                _members.Add(member.DeclaringType, new HashSet<string>());
            }
            
            _members[member.DeclaringType].Add(member.Name);
        }


        internal void ValidateType(MemberInfo member)
        {
            var isValid = false;
            foreach (var includedType in _includedTypes)
            {

                if (member.DeclaringType == includedType
                    || member.DeclaringType.IsAssignableFrom(includedType)
                    || member.DeclaringType.IsSubclassOf(includedType))
                {
                    isValid = true;
                    break;
                }
            }
            if (!isValid) throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
        }

        private List<ListType> CopyList<ListType>(List<ListType> source)
        {
            if (source == null) return null;
            var copy = new List<ListType>();
            source.ForEach(x => copy.Add(x));
            return copy;
        }

        private bool ShouldIgnoreMember(MemberInfo member, bool isNested)
        {
            if (member == null) return true;
            if (member.HasPropertyOfType<EpplusIgnore>()) return true;
            if (_members.Count == 0) return false;
            if (isNested && (_members == null || !_members.ContainsKey(member.DeclaringType)))
            {
                return false;
            }
            return !(_members.ContainsKey(member.DeclaringType) && _members[member.DeclaringType].Contains(member.Name));
        }

        private bool SetupInternal(Type type, List<ColumnInfo> result, List<int> sortOrderListArg, int nestedLevel = 0, string path = null, string headerPrefix = null)
        {
            var sort = false;
            var members = nestedLevel == 0 && _filterMembers != null ? _filterMembers.ToArray() : type.GetProperties(_bindingFlags);
            if (type.HasMemberWithPropertyOfType<EpplusTableColumnAttribute>() || type.HasMemberWithPropertyOfType<EpplusNestedTableColumnAttribute>())
            {
                sort = true;
                var index = 0;
                foreach (var member in members)
                {
                    if(member.DeclaringType != type && !member.DeclaringType.IsAssignableFrom(type)) continue;
                    var sortOrderList = CopyList(sortOrderListArg);
                    if (ShouldIgnoreMember(member, nestedLevel > 0))
                    {
                        continue;
                    }
                    if (member.HasPropertyOfType<EpplusIgnore>())
                    {
                        continue;
                    }
                    var memberPath = path != null ? $"{path}.{member.Name}" : member.Name;
                    if (member.HasPropertyOfType<EpplusNestedTableColumnAttribute>())
                    {
                        HandleNestedColumn(result, member, sortOrderList, headerPrefix, memberPath, nestedLevel);
                        continue;
                    }
                    if(member.HasPropertyOfType<EPPlusDictionaryColumnAttribute>())
                    {
                        HandleDictionaryColumnsAttribute(result, member, sortOrderList, headerPrefix, memberPath, ref index, nestedLevel);
                        continue;
                    }
                    //var sortOrder = GetSortOrder(memberPath);
                    _sortOrderCalculator.CalculateSortOrder(ref sortOrderList, index, nestedLevel, member);
                    var cs = TableColumnSettings.Default;
                    var colInfoSortOrderList = new List<int>();
                    var epplusColumnAttr = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
                    if (epplusColumnAttr != null)
                    {
                        /*
                         * private void HandleEpplusColumn(
            string headerPrefix, 
            List<int> sortOrderList, 
            TableColumnSettings colSettings,
            string memberPath, 
            List<int> colInfoSortOrderList, 
            EpplusTableColumnAttribute epplusColumnAttr, 
            int nestedLevel,
            int memberIndex,
            MemberInfo memberInfo)
                         */
                        HandleEpplusColumn(
                            headerPrefix, 
                            sortOrderList, 
                            ref cs, 
                            memberPath,
                            colInfoSortOrderList, 
                            epplusColumnAttr, 
                            nestedLevel,
                            index,
                            member);
                    }
                    else if(!string.IsNullOrEmpty(headerPrefix))
                    {
                        var header = string.IsNullOrEmpty(cs.Header) ? member.Name : cs.Header;
                        cs.Header = $"{headerPrefix} {header}";
                    }
                    else
                    {
                        cs.Header = string.IsNullOrEmpty(cs.Header) ? member.Name : cs.Header;
                    }
                    result.Add(new ColumnInfo
                    {
                        Header = cs.Header,
                        SortOrder = 0,
                        Index = index,
                        Hidden = cs.Hidden,
                        SortOrderLevels = colInfoSortOrderList,
                        MemberInfo = member,
                        NumberFormat = cs.NumberFormat,
                        TotalsRowFunction = cs.TotalsRowFunction,
                        TotalsRowNumberFormat = cs.TotalRowsNumberFormat,
                        TotalsRowLabel = cs.TotalRowLabel,
                        TotalsRowFormula = cs.TotalRowFormula,
                        Path = memberPath
                    });
                }
            }
            else
            {
                HandleNoExistingColumnAttributes(result, sortOrderListArg, nestedLevel, path, headerPrefix, members);
            }
            var formulaColumnAttributes = type.FindAttributesOfType<EpplusFormulaTableColumnAttribute>();
            if (formulaColumnAttributes != null && formulaColumnAttributes.Any())
            {
                sort = true;
                foreach (var attr in formulaColumnAttributes)
                {
                    result.Add(new ColumnInfo
                    {
                        SortOrder = attr.Order + SortOrderOffset,
                        Header = attr.Header,
                        Formula = attr.Formula,
                        FormulaR1C1 = attr.FormulaR1C1,
                        NumberFormat = attr.NumberFormat,
                        TotalsRowFunction = attr.TotalsRowFunction,
                        TotalsRowNumberFormat = attr.TotalsRowNumberFormat
                    });
                }
            }
            return sort;
        }

        private void HandleDictionaryColumnsAttribute(List<ColumnInfo> result, MemberInfo member, List<int> sortOrderList, string headerPrefix, string memberPath, ref int index, int nestedLevel)
        {
            var attr = member.GetFirstAttributeOfType<EPPlusDictionaryColumnAttribute>();
            if(member.MemberType == MemberTypes.Property)
            {
                if(((PropertyInfo)member).PropertyType != typeof(Dictionary<string, object>))
                {
                    throw new InvalidOperationException($"Property {memberPath} is decorated with the EPPlusDictionaryColumnsAttribute. Its type must be Dictionary<string, object>");
                }
            }
            else if (member.MemberType == MemberTypes.Field)
            {
                if (((FieldInfo)member).FieldType != typeof(Dictionary<string, object>))
                {
                    throw new InvalidOperationException($"Field {memberPath} is decorated with the EPPlusDictionaryColumnsAttribute. Its type must be Dictionary<string, object>");
                }
            }
            else if (member.MemberType == MemberTypes.Method)
            {
                if (((MethodInfo)member).ReturnType != typeof(Dictionary<string, object>))
                {
                    throw new InvalidOperationException($"Method {memberPath} is decorated with the EPPlusDictionaryColumnsAttribute. Its type must be Dictionary<string, object>");
                }
            }

            //var so = GetSortOrder(memberPath, attr.Order > -1 ? attr.Order : default);
            _sortOrderCalculator.CalculateSortOrder(ref sortOrderList, index, nestedLevel, member);
            var columnHeaders = Enumerable.Empty<string>();
            if(!string.IsNullOrEmpty(attr.KeyId))
            {
                columnHeaders = _params.GetDictionaryKeys(attr.KeyId);
            }
            else if(attr.ColumnHeaders != null && attr.ColumnHeaders.Length > 0)
            {
                columnHeaders = attr.ColumnHeaders;
            }
            else
            {
                columnHeaders = _params.GetDefaultDictionaryKeys();
            }
            foreach (var key in columnHeaders)
            {
                result.Add(new ColumnInfo
                {
                    Index = index++,
                    MemberInfo = member,
                    IsDictionaryProperty = true,
                    DictinaryKey = key,
                    Path = $"{memberPath}.{key}",
                    Header = key,
                    //SortOrder = so
                });
            }
        }

        private void HandleNoExistingColumnAttributes(List<ColumnInfo> result, List<int> sortOrderListArg, int nestedLevel, string path, string headerPrefix, MemberInfo[] members)
        {
            var index = 0;
            result.AddRange(members
                .Where(x => !x.HasPropertyOfType<EpplusIgnore>() && !ShouldIgnoreMember(x, nestedLevel > 0))
                .Select(member =>
                {
                    var h = default(string);
                    var mp = default(string);
                    if (!string.IsNullOrEmpty(path))
                    {
                        mp = $"{path}.{member.Name}";
                    }
                    var colInfoSortOrderList = new List<int>();
                    var sortOrderColumnsIndex = _sortOrderColumns != null ? _sortOrderColumns.IndexOf(mp) : -1;
                    var sortOrder = sortOrderColumnsIndex > -1 ? sortOrderColumnsIndex : nestedLevel * SortOrderOffset;
                    var sortOrderList = CopyList(sortOrderListArg);
                    var epplusColumnAttr = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
                    if (epplusColumnAttr != null)
                    {
                        h = string.IsNullOrEmpty(epplusColumnAttr.Header) ? member.Name : epplusColumnAttr.Header;
                        sortOrder = sortOrderColumnsIndex > -1 ? sortOrderColumnsIndex : epplusColumnAttr.Order + SortOrderOffset;
                    }
                    else
                    {
                        h = member.Name;
                    }

                    if (sortOrderList != null && sortOrderList.Any())
                    {
                        if (sortOrderColumnsIndex > -1)
                        {
                            sortOrderList[0] = sortOrder;
                        }
                        colInfoSortOrderList.AddRange(sortOrderList);
                    }

                    if (!string.IsNullOrEmpty(headerPrefix))
                    {
                        h = $"{headerPrefix} {h}";
                    }
                    else
                    {
                        h = member.Name;
                    }
                    return new ColumnInfo
                    {
                        Index = index++,
                        MemberInfo = member,
                        Path = mp,
                        Header = h,
                        SortOrder = sortOrder,
                        SortOrderLevels = colInfoSortOrderList
                    };
                }));
        }

        private void HandleEpplusColumn(
            string headerPrefix, 
            List<int> sortOrderList, 
            ref TableColumnSettings colSettings,
            string memberPath, 
            List<int> colInfoSortOrderList, 
            EpplusTableColumnAttribute epplusColumnAttr, 
            int nestedLevel,
            int memberIndex,
            MemberInfo memberInfo)
        {
            colSettings.SetProperties(epplusColumnAttr);
            if (!string.IsNullOrEmpty(epplusColumnAttr.Header) && !string.IsNullOrEmpty(headerPrefix))
            {
                colSettings.Header = $"{headerPrefix} {epplusColumnAttr.Header}";
            }
            else
            {
                colSettings.Header = epplusColumnAttr.Header;
            }
            //var sortOrder = GetSortOrder(memberPath, epplusColumnAttr.Order);
            _sortOrderCalculator.CalculateSortOrder(ref sortOrderList, memberIndex, nestedLevel, memberInfo);
            //sortOrder = sortOrderColumnsIndex > -1 ? sortOrderColumnsIndex : epplusColumnAttr.Order + nestedLevel * SortOrderOffset;
            _sortOrderCalculator.CalculateSortOrder(ref sortOrderList, memberIndex, nestedLevel, memberInfo);
            //if (sortOrderList != null && sortOrderList.Any())
            //{
            //    if (sortOrderColumnsIndex > -1)
            //    {
            //        sortOrderList[0] = sortOrder;
            //    }
            //    colInfoSortOrderList.AddRange(sortOrderList);
            //}
            //colInfoSortOrderList.Add(sortOrder < SortOrderOffset ? sortOrder : epplusColumnAttr.Order + nestedLevel * SortOrderOffset);
        }

        private void HandleNestedColumn(List<ColumnInfo> result, MemberInfo member, List<int> sortOrderList, string headerPrefix, string memberPath, int nestedLevel)
        {
            var hPrefix = default(string);
            var memberType = GetTypeByMemberInfo(member);
            if (memberType == typeof(string) || (!memberType.IsClass && memberType.IsInterface))
            {
                throw new InvalidOperationException($"EpplusNestedTableColumn attribute can only be used with complex types (member: {memberPath})");
            }
            var nestedTableAttr = member.GetFirstAttributeOfType<EpplusNestedTableColumnAttribute>();
            var attrOrder = nestedTableAttr.Order;
            // filter members override attribute order...
            if(_filterMembers != null && _filterMembers.Contains(member))
            {
                attrOrder = _filterMembers.IndexOf(member);
            }
            hPrefix = nestedTableAttr.HeaderPrefix;
            if (!string.IsNullOrEmpty(headerPrefix) && !string.IsNullOrEmpty(hPrefix))
            {
                hPrefix = $"{headerPrefix} {hPrefix}";
            }
            else if (!string.IsNullOrEmpty(headerPrefix))
            {
                hPrefix = headerPrefix;
            }
            if (_sortOrderColumns != null && _sortOrderColumns.IndexOf(memberPath) > -1)
            {
                attrOrder = _sortOrderColumns.IndexOf(memberPath);
            }
            else
            {
                attrOrder += SortOrderOffset;
            }
            //if (sortOrderList == null)
            //{
            //    sortOrderList = new List<int>
            //                {
            //                    attrOrder
            //                };
            //}
            //else
            //{
            //    sortOrderList.Add(attrOrder);
            //    if (attrOrder < SortOrderOffset)
            //    {
            //        sortOrderList[0] = _sortOrderColumns.IndexOf(memberPath);
            //    }
            //}
            SetupInternal(memberType, result, sortOrderList, nestedLevel + 1, memberPath, hPrefix);
            sortOrderList.RemoveAt(sortOrderList.Count - 1);
        }

        private Type GetTypeByMemberInfo(MemberInfo member)
        {
            switch(member.MemberType)
            {
                case MemberTypes.Field:
                    return ((FieldInfo)member).FieldType;
                case MemberTypes.Property:
                    return ((PropertyInfo)member).PropertyType;
                case MemberTypes.Method:
                    return ((MethodInfo)member).ReturnType;
                default:
                    throw new InvalidOperationException($"LoadFromCollection: Unsupported MemberType on member {member.Name}. Only Field, Property and Method allowed.");
            }
        }

        private static void ReindexAndSortColumns(List<ColumnInfo> result)
        {
            var index = 0;
            //result.Sort((a, b) => a.SortOrder.CompareTo(b.SortOrder));
            result.Sort((a, b) =>
            {
                var so1 = a.SortOrderLevels;
                var so2 = b.SortOrderLevels;
                if (so1 == null || so2 == null)
                {
                    if(a.SortOrder == b.SortOrder)
                    {
                        return a.Index.CompareTo(b.Index);
                    }
                    else
                    {
                        return a.SortOrder.CompareTo(b.SortOrder);
                    }
                }
                else if (!so1.Any() || !so2.Any())
                {
                    return a.SortOrder.CompareTo(b.SortOrder);
                }
                else
                {
                    var maxIx = so1.Count < so2.Count ? so1.Count : so2.Count;
                    for(var ix = 0; ix < maxIx; ix++)
                    {
                        var aVal = so1[ix];
                        var bVal = so2[ix];
                        if (aVal.CompareTo(bVal) == 0) continue;
                        return aVal.CompareTo(bVal);
                    }
                    return a.Index.CompareTo(b.Index);
                }
            });
            result.ForEach(x => x.Index = index++);
        }        
    }
}
