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
    internal class LoadFromCollectionColumns_old<T>
    {
        //public LoadFromCollectionColumns(LoadFromCollectionParams parameters) :
        //    this(parameters, Enumerable.Empty<string>().ToList())
        //{ }

        public LoadFromCollectionColumns_old(LoadFromCollectionParams parameters)
        {
            _params = parameters;
            //_membersStore = new MembersStore<T>(parameters.Members, parameters.BindingFlags);
            //var typeScanner = _membersStore.GetTypeScanner();
            //_sortOrderCalculator = new SortOrderCalculator(typeScanner, parameters.Members);
            //_bindingFlags = parameters.BindingFlags;
            //_sortOrderColumns = sortOrderColumns;
            //if (parameters.Members != null)
            //{
            //    _filterMembers = parameters.Members.ToList();
            //}
            //var typeScanner = new NestedColumnsTypeScanner(typeof(T), _bindingFlags);
            //_sortOrderCalculator = new SortOrderCalculator(typeScanner, parameters.Members);
            //var members = parameters.Members;
            //_members = new Dictionary<Type, HashSet<string>>();
            //if (members != null && members.Length > 0)
            //{
            //    foreach (var member in members)
            //    {
            //        AddMember(member);
            //    }
            //}
        }


        private readonly LoadFromCollectionParams _params;
        private readonly MembersStore<T> _membersStore;
        private readonly BindingFlags _bindingFlags;
        private readonly List<string> _sortOrderColumns;
        private readonly SortOrderCalculator _sortOrderCalculator;
        private const int SortOrderOffset = ExcelPackage.MaxColumns;

        internal ColumnInfoCollection Setup()
        {
            var result = new ColumnInfoCollection();
            var t = typeof(T);
            var ut = Nullable.GetUnderlyingType(t);
            if (ut != null)
            {
                t = ut;
            }

            SetupInternal(t, result, null);
            result.ReindexAndSortColumns();
            return result;
        }


        internal void ValidateType(MemberInfo member)
        {
            _membersStore.ValidateType(member);
        }

        private List<ListType> CopyList<ListType>(List<ListType> source)
        {
            if (source == null) return null;
            var copy = new List<ListType>();
            source.ForEach(x => copy.Add(x));
            return copy;
        }

        

        private void SetupInternal(
            Type type, 
            List<ColumnInfo> result, 
            List<int> sortOrderListArg, 
            int nestedLevel = 0, 
            MemberPath path = null, 
            string headerPrefix = null)
        {
            var members = _membersStore.GetMembers(type, nestedLevel);
            var index = 0;
            foreach (var member in members)
            {
                if (member.DeclaringType != type && !member.DeclaringType.IsAssignableFrom(type)) continue;
                var sortOrderList = CopyList(sortOrderListArg);
                if(path == null)
                {
                    path = new MemberPath(member, 0);
                }
                else
                {
                    path.Append(member, 0);
                }
                if (_membersStore.ShouldIgnoreMember(member, path))
                {
                    continue;
                }
                var cs = TableColumnSettings.Default;
                //var memberPath = path != null ? $"{path}.{member.Name}" : member.Name;
                _sortOrderCalculator.CalculateSortOrder(ref sortOrderList, index, nestedLevel, path, member);
                if (member.HasAttributeOfType(out EpplusNestedTableColumnAttribute nestedTableAttr))
                {
                    var memberType = MemberHelper.GetTypeByMemberInfo(member);
                    if (memberType == typeof(string) || (!memberType.IsClass && memberType.IsInterface))
                    {
                        throw new InvalidOperationException($"EpplusNestedTableColumn attribute can only be used with complex types (member: {path.GetPath()})");
                    }
                    var hPrefix = ColumnsHeaderReader.GetAggregatedHeaderPrefix(headerPrefix, nestedTableAttr);
                    //_sortOrderCalculator.CalculateSortOrder(ref sortOrderList, index, nestedLevel, member);
                    SetupInternal(memberType, result, sortOrderList, nestedLevel + 1, path, hPrefix);
                    index++;
                    continue;
                }
                else if (member.HasAttributeOfType<EPPlusDictionaryColumnAttribute>())
                {
                    //_sortOrderCalculator.CalculateSortOrder(ref sortOrderList, index, nestedLevel, member);
                    HandleDictionaryColumnsAttribute(result, member, sortOrderList, headerPrefix, path, index++, nestedLevel);
                    continue;
                }
                else if (member.HasAttributeOfType(out EpplusTableColumnAttribute epplusColumnAttr))
                {
                    cs.SetProperties(epplusColumnAttr);
                }
                else
                {
                    //_sortOrderCalculator.CalculateSortOrder(ref sortOrderList, index, nestedLevel, member);
                }

                result.Add(new ColumnInfo
                {
                    Header = cs.GetHeader(headerPrefix, member),
                    Index = index,
                    Hidden = cs.Hidden,
                    //SortOrderLevels = sortOrderList,
                    MemberInfo = member,
                    NumberFormat = cs.NumberFormat,
                    TotalsRowFunction = cs.TotalsRowFunction,
                    TotalsRowNumberFormat = cs.TotalRowsNumberFormat,
                    TotalsRowLabel = cs.TotalRowLabel,
                    TotalsRowFormula = cs.TotalRowFormula,
                    //Path = path.GetPath()
                });
                index++;
            }
            var formulaColumnAttributes = type.FindAttributesOfType<EpplusFormulaTableColumnAttribute>();
            if (formulaColumnAttributes != null && formulaColumnAttributes.Any())
            {
                foreach (var attr in formulaColumnAttributes)
                {
                    result.Add(new ColumnInfo
                    {
                        //SortOrderLevels = new List<int> { attr.Order },
                        Header = attr.Header,
                        Formula = attr.Formula,
                        FormulaR1C1 = attr.FormulaR1C1,
                        NumberFormat = attr.NumberFormat,
                        TotalsRowFunction = attr.TotalsRowFunction,
                        TotalsRowNumberFormat = attr.TotalsRowNumberFormat
                    });
                }
            }
        }

        private void HandleDictionaryColumnsAttribute(List<ColumnInfo> result, MemberInfo member, List<int> sortOrderList, string headerPrefix, MemberPath memberPath, int index, int nestedLevel)
        {
            var attr = member.GetFirstAttributeOfType<EPPlusDictionaryColumnAttribute>();
            if (member.MemberType == MemberTypes.Property)
            {
                if (((PropertyInfo)member).PropertyType != typeof(Dictionary<string, object>))
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
            var columnHeaders = Enumerable.Empty<string>();
            if (!string.IsNullOrEmpty(attr.KeyId))
            {
                columnHeaders = _params.GetDictionaryKeys(attr.KeyId);
            }
            else if (attr.ColumnHeaders != null && attr.ColumnHeaders.Length > 0)
            {
                columnHeaders = attr.ColumnHeaders;
            }
            else
            {
                columnHeaders = _params.GetDefaultDictionaryKeys();
            }
            foreach (var key in columnHeaders)
            {
                var sortOrderListCol = CopyList(sortOrderList);
                _sortOrderCalculator.CalculateSortOrder(ref sortOrderListCol, index++, nestedLevel, memberPath, member);
                result.Add(new ColumnInfo
                {
                    Index = index,
                    MemberInfo = member,
                    IsDictionaryProperty = true,
                    DictinaryKey = key,
                    //Path = $"{memberPath}.{key}",
                    Header = key,
                    //SortOrderLevels = sortOrderListCol
                });
            }
        }
    }
}
