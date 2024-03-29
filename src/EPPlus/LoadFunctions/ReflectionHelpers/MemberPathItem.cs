﻿using OfficeOpenXml.Table;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions.ReflectionHelpers
{
    [DebuggerDisplay("Member: {Member.Name}, SortOrder: {SortOrder}")]
    internal class MemberPathItem
    {
        public MemberPathItem(EpplusFormulaTableColumnAttribute attr)
        {
            SortOrder = attr.Order;
            NumberFormat = attr.NumberFormat;
            TotalsRowFunction = attr.TotalsRowFunction;
            TotalRowsNumberFormat = attr.TotalsRowNumberFormat;
            TotalRowLabel = attr.TotalsRowLabel;
            TotalRowFormula = attr.TotalsRowFormula;
        }

        public MemberPathItem(MemberInfo member, string dictionaryKey, int index)
        {
            Member = member;
            DictionaryKey = dictionaryKey;
            IsDictionaryColumn = true;
            SortOrder = index;
        }

        public MemberPathItem(MemberInfo member, int sortOrder)
        {
            Member = member;
            SortOrder = sortOrder;
        }

        public MemberInfo Member { get; set; }

        public MemberPathItem Parent { get; set; }

        public bool IsNestedProperty { get; set; }

        public string HeaderPrefix { get; set; }

        public int SortOrder { get; set; }

        public bool IsDictionaryColumn { get; set; }

        public bool IsDictionaryParent { get; set; }

        public string DictionaryKey { get; set; }

        public bool Hidden { get; set; }

        public string NumberFormat { get; set; }

        public RowFunctions TotalsRowFunction { get; set; }

        public string TotalRowsNumberFormat { get; set; }

        public string TotalRowLabel { get; set; }

        public string TotalRowFormula { get; set; }


        public MemberPathItem Clone()
        {
            return new MemberPathItem(Member, SortOrder)
            {
                Parent = Parent,
                IsNestedProperty = IsNestedProperty,
                HeaderPrefix = HeaderPrefix,
                IsDictionaryColumn = IsDictionaryColumn,
                Hidden = Hidden,
                NumberFormat = NumberFormat,
                TotalsRowFunction = TotalsRowFunction,
                TotalRowsNumberFormat = TotalRowsNumberFormat,
                TotalRowLabel = TotalRowLabel,
                TotalRowFormula = TotalRowFormula,
            };
        }

        public void SetProperties(EpplusTableColumnAttribute attr)
        {
            Hidden = attr.Hidden;
            NumberFormat = attr.NumberFormat;
            TotalsRowFunction = attr.TotalsRowFunction;
            TotalRowsNumberFormat = attr.TotalsRowNumberFormat;
            TotalRowLabel = attr.TotalsRowLabel;
            TotalRowFormula = attr.TotalsRowFormula;
        }

        public void SetProperties(EpplusNestedTableColumnAttribute attr)
        {
            IsNestedProperty = true;
            HeaderPrefix = attr.HeaderPrefix;
        }

        private Dictionary<Type, Func<object, object>> _lambdas = new Dictionary<Type, Func<object, object>>();

        private void CreatePropertyGetter(Type outerType)
        {
            if(outerType.GetProperty(Member.Name) != null)
            {
                var parameter = Expression.Parameter(typeof(object), "x");
                Expression expression = Expression.Convert(parameter, outerType);

                expression = Expression.Property(expression, Member.Name);

                var lambda = Expression.Lambda<Func<object, object>>(Expression.Convert(expression, typeof(object)), parameter);
                _lambdas[outerType] = lambda.Compile();
            }
            else
            {
                _lambdas[outerType] = null;
            }
        }

        public object GetPropertyValue(object item)
        {
            var type = item.GetType();
            if(_lambdas.ContainsKey(type))
            {
                var l = _lambdas[type];
                if(l != null)
                {
                    return l(item);
                }
                return null;
            }
            else
            {
                CreatePropertyGetter(type);
                var l = _lambdas[type];
                if (l != null)
                {
                    return l(item);
                }
                return null;
            }
        }
    }
}
